"""
Transport layer for the PyExcel kernel — the wire under :mod:`framing`.

The Python kernel is a client: it is spawned by the C# supervisor, which has
already created a named-pipe server. We dial in, then framing reads/writes
flow over the resulting bidirectional byte stream.

Two backends:

* **POSIX** (Linux/macOS — what CI runs). .NET's ``NamedPipeServerStream``
  on POSIX maps the pipe to a Unix domain socket at
  ``$TMPDIR/CoreFxPipe_<name>`` (default ``/tmp``). We connect with
  stdlib ``socket(AF_UNIX, SOCK_STREAM)`` — no third-party deps.
* **Windows** — production. Uses stdlib ``_winapi`` to call
  ``CreateFile`` / ``WaitNamedPipe`` / ``ReadFile`` / ``WriteFile``
  against ``\\\\.\\pipe\\<name>``. The C# server sets a DACL at pipe
  creation that allows only the current-user SID, so wrong-user
  processes get ERROR_ACCESS_DENIED at connect time before any bytes
  cross the boundary.

The returned :class:`FrameTransport` exposes ``read_exact`` (the callable
shape :func:`framing.read_frame` consumes) and ``write_all`` for sending
encoded frames. It is intentionally synchronous and blocking — the
supervisor is a single-threaded event loop.
"""

from __future__ import annotations

import os
import socket
import sys
import tempfile
import time
from typing import Optional


class TransportError(Exception):
    """Anything that prevents a frame from completing on the wire."""


class FrameTransport:
    """Thin bidirectional byte-stream wrapper used by the framing layer.

    Subclasses implement the two primitives; everything else (close, context
    manager) lives here.
    """

    def read_exact(self, n: int) -> bytes:  # pragma: no cover - abstract
        raise NotImplementedError

    def write_all(self, data: bytes) -> None:  # pragma: no cover - abstract
        raise NotImplementedError

    def close(self) -> None:  # pragma: no cover - abstract
        raise NotImplementedError

    def __enter__(self) -> "FrameTransport":
        return self

    def __exit__(self, *exc_info) -> None:
        self.close()


class _SocketTransport(FrameTransport):
    """``AF_UNIX`` (POSIX) implementation backing a .NET pipe-on-Linux."""

    def __init__(self, sock: socket.socket) -> None:
        self._sock = sock

    def read_exact(self, n: int) -> bytes:
        # recv_into would save one allocation per call but socket.recv is fine
        # for the frame cadence we care about (request/response, not stream).
        buf = bytearray()
        while len(buf) < n:
            chunk = self._sock.recv(n - len(buf))
            if not chunk:
                # Short return on a SOCK_STREAM means the peer closed; framing
                # interprets a short read as EOF and raises TruncatedFrameError.
                return bytes(buf)
            buf.extend(chunk)
        return bytes(buf)

    def write_all(self, data: bytes) -> None:
        try:
            self._sock.sendall(data)
        except OSError as exc:
            raise TransportError(f"socket write failed: {exc}") from exc

    def close(self) -> None:
        try:
            self._sock.shutdown(socket.SHUT_RDWR)
        except OSError:
            pass
        self._sock.close()


def _posix_pipe_path(pipe_name: str) -> str:
    # Matches the prefix .NET's PipeStream.Unix.cs uses on POSIX, so any
    # NamedPipeServerStream created in C# is reachable from here.
    return os.path.join(tempfile.gettempdir(), f"CoreFxPipe_{pipe_name}")


def _connect_posix(
    pipe_name: str,
    connect_timeout_s: float,
) -> _SocketTransport:
    path = _posix_pipe_path(pipe_name)
    deadline = time.monotonic() + connect_timeout_s
    last_exc: Optional[OSError] = None

    # The server may not have called WaitForConnection yet when we wake. Retry
    # ENOENT/ECONNREFUSED until the deadline before giving up.
    while True:
        sock = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        try:
            sock.connect(path)
            return _SocketTransport(sock)
        except OSError as exc:
            sock.close()
            last_exc = exc
            if time.monotonic() >= deadline:
                raise TransportError(
                    f"could not connect to pipe {path!r} within "
                    f"{connect_timeout_s}s: {exc}"
                ) from exc
            time.sleep(0.05)


class _WinPipeTransport(FrameTransport):
    """Win32 named-pipe client implementation."""

    # Win32 error codes we treat as "peer closed" rather than a hard failure.
    # 109 (ERROR_BROKEN_PIPE) and 233 (ERROR_PIPE_NOT_CONNECTED) both mean
    # the server end is gone; framing turns the resulting short read into
    # TruncatedFrameError.
    _EOF_ERRORS = frozenset({109, 233})

    def __init__(self, handle) -> None:
        self._handle = handle
        self._closed = False

    def read_exact(self, n: int) -> bytes:
        if n == 0:
            return b""
        import _winapi  # type: ignore[attr-defined]
        buf = bytearray()
        while len(buf) < n:
            try:
                chunk, _err = _winapi.ReadFile(self._handle, n - len(buf), False)
            except OSError as exc:
                if getattr(exc, "winerror", 0) in self._EOF_ERRORS:
                    return bytes(buf)
                raise TransportError(f"pipe read failed: {exc}") from exc
            if not chunk:
                # Server closed cleanly between our calls.
                return bytes(buf)
            buf.extend(chunk)
        return bytes(buf)

    def write_all(self, data: bytes) -> None:
        if not data:
            return
        import _winapi  # type: ignore[attr-defined]
        mv = memoryview(data)
        pos = 0
        while pos < len(mv):
            try:
                written, _err = _winapi.WriteFile(self._handle, mv[pos:], False)
            except OSError as exc:
                raise TransportError(f"pipe write failed: {exc}") from exc
            if written == 0:
                # Defensive: WriteFile returning 0 with no error would be a
                # bug in the underlying API, but we'd rather raise than spin.
                raise TransportError("pipe write returned 0 bytes")
            pos += written

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        try:
            import _winapi  # type: ignore[attr-defined]
            _winapi.CloseHandle(self._handle)
        except OSError:
            # Handle already gone (e.g. peer closed and OS reaped it).
            pass


def _connect_windows(
    pipe_name: str,
    connect_timeout_s: float,
) -> FrameTransport:
    # .NET's ``NamedPipeServerStream`` exposes the pipe at the canonical
    # Win32 path ``\\.\pipe\<name>`` — no platform-specific prefix mapping
    # like the POSIX side has. We open it with read/write access using the
    # stdlib ``_winapi`` shim (the same module Python's ``multiprocessing``
    # uses for its own pipe transport).
    import _winapi  # type: ignore[attr-defined]

    # The Win32 path is \\.\pipe\<name> — one backslash before the name.
    # Building this with a raw string would need a trailing single backslash,
    # which Python raw-string syntax forbids; concatenate the separator
    # explicitly to avoid the ambiguity.
    path = r"\\.\pipe" + "\\" + pipe_name
    deadline = time.monotonic() + connect_timeout_s

    # ERROR_PIPE_BUSY: server saw a connection on its single instance but
    # hasn't called WaitForConnection again yet. WaitNamedPipe blocks until
    # a slot is free or the deadline lapses; on ERROR_FILE_NOT_FOUND the
    # server hasn't created the pipe yet, so we just retry with a short
    # sleep.
    ERROR_FILE_NOT_FOUND = 2
    ERROR_PIPE_BUSY = 231

    last_exc: Optional[OSError] = None
    while True:
        try:
            handle = _winapi.CreateFile(
                path,
                _winapi.GENERIC_READ | _winapi.GENERIC_WRITE,
                0,  # exclusive — no other handles to the same pipe
                _winapi.NULL,
                _winapi.OPEN_EXISTING,
                0,  # synchronous I/O; the supervisor is a single-threaded loop
                _winapi.NULL,
            )
            return _WinPipeTransport(handle)
        except OSError as exc:
            last_exc = exc
            winerror = getattr(exc, "winerror", 0)
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise TransportError(
                    f"could not connect to pipe {path!r} within "
                    f"{connect_timeout_s}s: {exc}"
                ) from exc

            if winerror == ERROR_PIPE_BUSY:
                # Block up to the remaining budget waiting for a free slot.
                wait_ms = max(1, int(remaining * 1000))
                try:
                    _winapi.WaitNamedPipe(path, wait_ms)
                except OSError:
                    # Fall through to the retry; the next CreateFile will
                    # surface the underlying reason in its own OSError.
                    pass
                continue

            if winerror == ERROR_FILE_NOT_FOUND:
                # Server hasn't created the pipe yet. Short sleep + retry.
                time.sleep(0.05)
                continue

            # Anything else (ERROR_ACCESS_DENIED from the DACL, etc.) is a
            # terminal failure — no point retrying.
            raise TransportError(
                f"could not connect to pipe {path!r}: {exc}"
            ) from exc


def connect(
    pipe_name: str,
    *,
    connect_timeout_s: float = 5.0,
) -> FrameTransport:
    """Connect to the supervisor's pipe and return a ready transport.

    Raises :class:`TransportError` on dial failure (timeout or refused).
    """
    if not pipe_name or not isinstance(pipe_name, str):
        raise ValueError("pipe_name must be a non-empty string")

    if sys.platform == "win32":
        return _connect_windows(pipe_name, connect_timeout_s)
    return _connect_posix(pipe_name, connect_timeout_s)
