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
* **Windows** — production. Uses the public Win32 named-pipe API via
  ``ctypes`` (stdlib). Implemented but not exercised by CI today; the
  Windows .NET tests don't yet spawn a kernel.

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


def _connect_windows(
    pipe_name: str,
    connect_timeout_s: float,
) -> FrameTransport:
    # Deferred until the Windows kernel test slice is added. The .NET side
    # already opens the pipe at \\.\pipe\<name>; the equivalent client call
    # is CreateFileW with GENERIC_READ | GENERIC_WRITE. Wiring it up needs
    # ctypes signatures for ReadFile/WriteFile/CloseHandle plus the
    # WaitNamedPipe retry. Tracked alongside the Windows integration test.
    raise NotImplementedError(
        "windows named-pipe transport not implemented yet; "
        "the linux/posix path is the one CI exercises today"
    )


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
