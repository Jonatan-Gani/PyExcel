"""End-to-end test for the kernel supervisor loop.

We act as the C# side — create a Unix-domain socket where .NET's
``NamedPipeServerStream`` would put it on POSIX
(``$TMPDIR/CoreFxPipe_<name>``), spawn ``python -m pyexcel.kernel`` to dial
in, then drive the HELLO / PING / SHUTDOWN protocol from this side. This
gives us confidence in the Python half independent of the C# integration
test, which exercises the same path from the other direction.

Skipped on Windows: this test plays the C# *server* role from Python, which
would require reimplementing the server side of the named-pipe protocol
via ``_winapi.CreateNamedPipe``. The C# integration tests (under
``tests/PyExcel.Bridge.Tests``) cover the Windows side end-to-end by
running the real C# server against the Python client.
"""

from __future__ import annotations

import os
import socket
import subprocess
import sys
import tempfile
import textwrap
import threading
import time
import uuid
from pathlib import Path

import pytest

from pyexcel.kernel import arrow_io
from pyexcel.kernel.framing import (
    PROTOCOL_VERSION,
    FrameType,
    encode_frame,
    read_frame,
)

pytestmark = pytest.mark.skipif(
    sys.platform == "win32",
    reason="this test acts as the C# server; reimplementing that side from "
    "Python on Windows is out of scope (covered by the C# integration tests)",
)


REPO_ROOT = Path(__file__).resolve().parents[2]
EMBEDDED = REPO_ROOT / "embedded"


def _pipe_path(pipe_name: str) -> str:
    return os.path.join(tempfile.gettempdir(), f"CoreFxPipe_{pipe_name}")


class _Server:
    """Minimal AF_UNIX server that pretends to be C# KernelSupervisor."""

    def __init__(self, pipe_name: str) -> None:
        self.pipe_name = pipe_name
        self.path = _pipe_path(pipe_name)
        self._listener = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        # Clean up any stale socket file from an aborted previous run.
        try:
            os.unlink(self.path)
        except FileNotFoundError:
            pass
        self._listener.bind(self.path)
        self._listener.listen(1)
        self._conn: socket.socket | None = None

    def accept(self, timeout_s: float = 5.0) -> None:
        self._listener.settimeout(timeout_s)
        self._conn, _ = self._listener.accept()
        self._conn.settimeout(timeout_s)

    def send(self, frame_type: FrameType, meta: dict) -> None:
        assert self._conn is not None
        self._conn.sendall(encode_frame(frame_type, meta))

    def recv(self):
        assert self._conn is not None

        def read_exact(n: int) -> bytes:
            buf = bytearray()
            while len(buf) < n:
                chunk = self._conn.recv(n - len(buf))
                if not chunk:
                    return bytes(buf)
                buf.extend(chunk)
            return bytes(buf)

        return read_frame(read_exact)

    def close(self) -> None:
        for s in (self._conn, self._listener):
            if s is None:
                continue
            try:
                s.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
            s.close()
        try:
            os.unlink(self.path)
        except FileNotFoundError:
            pass


def _spawn_kernel(pipe_name: str) -> subprocess.Popen:
    env = os.environ.copy()
    env["PYTHONPATH"] = str(EMBEDDED) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONUNBUFFERED"] = "1"
    return subprocess.Popen(
        [sys.executable, "-X", "utf8", "-m", "pyexcel.kernel", "--pipe", pipe_name],
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )


def _drain(proc: subprocess.Popen, timeout_s: float = 5.0) -> tuple[int, str, str]:
    try:
        out, err = proc.communicate(timeout=timeout_s)
    except subprocess.TimeoutExpired:
        proc.kill()
        out, err = proc.communicate()
    return proc.returncode, out.decode(errors="replace"), err.decode(errors="replace")


def test_kernel_handshake_ping_shutdown_roundtrip():
    pipe_name = "pyexcel-test-" + uuid.uuid4().hex
    server = _Server(pipe_name)
    proc = _spawn_kernel(pipe_name)

    try:
        server.accept(timeout_s=5.0)

        # Handshake: server (us) sends HELLO first.
        server.send(FrameType.HELLO, {"protocol": PROTOCOL_VERSION})
        client_hello = server.recv()
        assert client_hello.type is FrameType.HELLO
        assert client_hello.meta == {"protocol": PROTOCOL_VERSION}

        # PING/PONG with nonce echo.
        server.send(FrameType.PING, {"nonce": "n1"})
        pong = server.recv()
        assert pong.type is FrameType.PONG
        assert pong.meta == {"nonce": "n1"}

        # SHUTDOWN -> clean exit (return code 0).
        server.send(FrameType.SHUTDOWN, {})
        rc, _, err = _drain(proc, timeout_s=5.0)
        assert rc == 0, f"kernel exited with {rc}; stderr={err!r}"
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=2)
        server.close()


def test_kernel_rejects_wrong_protocol_in_hello():
    pipe_name = "pyexcel-test-" + uuid.uuid4().hex
    server = _Server(pipe_name)
    proc = _spawn_kernel(pipe_name)

    try:
        server.accept(timeout_s=5.0)
        server.send(FrameType.HELLO, {"protocol": PROTOCOL_VERSION + 999})

        # Kernel should reply with ERROR then exit non-zero. The order
        # depends on how fast the OS flushes — we only assert end state.
        rc, _, err = _drain(proc, timeout_s=5.0)
        assert rc != 0, "kernel should exit non-zero on protocol mismatch"
        assert "protocol mismatch" in err.lower() or "handshake" in err.lower()
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=2)
        server.close()


def test_kernel_run_request_with_bad_meta_returns_error_and_keeps_running():
    pipe_name = "pyexcel-test-" + uuid.uuid4().hex
    server = _Server(pipe_name)
    proc = _spawn_kernel(pipe_name)

    try:
        server.accept(timeout_s=5.0)
        server.send(FrameType.HELLO, {"protocol": PROTOCOL_VERSION})
        _ = server.recv()  # client HELLO

        # RUN_REQUEST missing 'script' → worker returns ERROR/BadRequest.
        server.send(FrameType.RUN_REQUEST, {"run_id": "bad"})
        err_frame = server.recv()
        assert err_frame.type is FrameType.ERROR
        assert err_frame.meta["code"] == "BadRequest"
        assert err_frame.meta["run_id"] == "bad"

        # Loop is intact: a PING still gets a PONG.
        server.send(FrameType.PING, {"nonce": "after-error"})
        pong = server.recv()
        assert pong.type is FrameType.PONG
        assert pong.meta == {"nonce": "after-error"}

        server.send(FrameType.SHUTDOWN, {})
        rc, _, err = _drain(proc, timeout_s=5.0)
        assert rc == 0, f"kernel exited with {rc}; stderr={err!r}"
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=2)
        server.close()


def test_kernel_run_request_executes_user_script_end_to_end(tmp_path):
    # Write a real .py script and drive the kernel through one full
    # RUN_REQUEST -> RUN_RESULT exchange. This is the integration sibling
    # of the unit tests in test_worker.py.
    script = tmp_path / "double.py"
    script.write_text(textwrap.dedent("""
        def transform(x):
            return x * 2
    """))

    pipe_name = "pyexcel-test-" + uuid.uuid4().hex
    server = _Server(pipe_name)
    proc = _spawn_kernel(pipe_name)

    try:
        server.accept(timeout_s=5.0)
        server.send(FrameType.HELLO, {"protocol": PROTOCOL_VERSION})
        _ = server.recv()  # client HELLO

        # Send the job: one Arrow-encoded scalar argument.
        payload = arrow_io.encode(21)
        encoded = encode_frame(
            FrameType.RUN_REQUEST,
            {"run_id": "job-1", "script": str(script)},
            (payload,),
        )
        server._conn.sendall(encoded)  # type: ignore[union-attr]

        result = server.recv()
        assert result.type is FrameType.RUN_RESULT, (
            f"got {result.type.name}; meta={result.meta!r}"
        )
        assert result.meta["run_id"] == "job-1"
        assert "duration_ms" in result.meta
        assert len(result.payloads) == 1
        assert arrow_io.decode(result.payloads[0]) == 42

        server.send(FrameType.SHUTDOWN, {})
        rc, _, err = _drain(proc, timeout_s=5.0)
        assert rc == 0, f"kernel exited with {rc}; stderr={err!r}"
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=2)
        server.close()


def test_kernel_run_request_user_exception_produces_error_frame(tmp_path):
    script = tmp_path / "boom.py"
    script.write_text(textwrap.dedent("""
        def transform(x):
            raise ValueError(f"no good: {x}")
    """))

    pipe_name = "pyexcel-test-" + uuid.uuid4().hex
    server = _Server(pipe_name)
    proc = _spawn_kernel(pipe_name)

    try:
        server.accept(timeout_s=5.0)
        server.send(FrameType.HELLO, {"protocol": PROTOCOL_VERSION})
        _ = server.recv()

        encoded = encode_frame(
            FrameType.RUN_REQUEST,
            {"run_id": "job-2", "script": str(script)},
            (arrow_io.encode(99),),
        )
        server._conn.sendall(encoded)  # type: ignore[union-attr]

        err_frame = server.recv()
        assert err_frame.type is FrameType.ERROR
        assert err_frame.meta["code"] == "Exception"
        assert err_frame.meta["type"] == "ValueError"
        assert "99" in err_frame.meta["message"]
        assert "boom.py" in err_frame.meta["traceback"]

        server.send(FrameType.SHUTDOWN, {})
        rc, _, _ = _drain(proc, timeout_s=5.0)
        assert rc == 0
    finally:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=2)
        server.close()
