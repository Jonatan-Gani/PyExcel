"""
Kernel-side supervisor loop.

Lifecycle of a kernel process:

    1. ``__main__`` parses argv, opens the transport, calls :func:`run`.
    2. :func:`run` completes the HELLO handshake against the C# side.
    3. The main loop reads one frame at a time, dispatches, replies
       inline. SHUTDOWN exits cleanly with code 0; anything that goes
       wrong on the wire exits 1 so the supervisor on the other end can
       detect the death and re-spawn.

Frames handled in-loop:

* ``HELLO`` — handshake (server speaks first, we mirror back).
* ``PING`` — answered with ``PONG`` echoing any ``nonce`` meta.
* ``RUN_REQUEST`` — delegated to :mod:`pyexcel.kernel.worker` and answered
  inline with either ``RUN_RESULT`` or ``ERROR`` depending on outcome.
* ``SHUTDOWN`` — clean exit with code 0.

Anything else gets an ``ERROR`` reply and the loop stays alive.
"""

from __future__ import annotations

import sys
from typing import NoReturn

from . import worker
from .framing import (
    PROTOCOL_VERSION,
    Frame,
    FrameType,
    FramingError,
    encode_frame,
    read_frame,
)
from .transport import FrameTransport, TransportError, connect


def _send(
    transport: FrameTransport,
    frame_type: FrameType,
    meta: dict,
    payloads: tuple = (),
) -> None:
    transport.write_all(encode_frame(frame_type, meta, payloads))


def _recv(transport: FrameTransport) -> Frame:
    return read_frame(transport.read_exact)


def _handshake(transport: FrameTransport) -> int:
    """Wait for the server's HELLO, verify protocol, reply with our own.

    Returns the negotiated protocol version (currently always the single
    supported value; once we have more than one we'll pick max(min(...))).
    """
    hello = _recv(transport)
    if hello.type is not FrameType.HELLO:
        _send(
            transport,
            FrameType.ERROR,
            {"stage": "handshake", "reason": f"expected HELLO, got {hello.type.name}"},
        )
        raise RuntimeError(f"handshake: expected HELLO, got {hello.type.name}")

    server_proto = hello.meta.get("protocol")
    if server_proto != PROTOCOL_VERSION:
        _send(
            transport,
            FrameType.ERROR,
            {
                "stage": "handshake",
                "reason": "protocol mismatch",
                "client": PROTOCOL_VERSION,
                "server": server_proto,
            },
        )
        raise RuntimeError(
            f"handshake: protocol mismatch (client={PROTOCOL_VERSION}, server={server_proto})"
        )

    _send(transport, FrameType.HELLO, {"protocol": PROTOCOL_VERSION})
    return PROTOCOL_VERSION


def _dispatch(transport: FrameTransport, frame: Frame) -> bool:
    """Handle one inbound frame. Return False to break the loop (SHUTDOWN)."""
    if frame.type is FrameType.PING:
        # Echo the nonce so the caller can match PONGs to PINGs even if some
        # other frame interleaves later. Empty meta is allowed for callers
        # that don't care.
        nonce = frame.meta.get("nonce", "")
        _send(transport, FrameType.PONG, {"nonce": nonce})
        return True

    if frame.type is FrameType.RUN_REQUEST:
        outcome = worker.run_job(frame.meta, frame.payloads)
        reply_type = FrameType.RUN_RESULT if outcome.success else FrameType.ERROR
        _send(transport, reply_type, outcome.meta, tuple(outcome.payloads))
        return True

    if frame.type is FrameType.SHUTDOWN:
        return False

    # Anything not in the dispatch table above (CANCEL, LIST_JOBS, …) gets a
    # polite ERROR rather than silent drop. The C# side surfaces this in tests.
    _send(
        transport,
        FrameType.ERROR,
        {"reason": f"unsupported frame type {frame.type.name} at this phase"},
    )
    return True


def run(pipe_name: str, *, connect_timeout_s: float = 5.0) -> int:
    """Connect, handshake, and run the main loop. Returns the process exit code."""
    try:
        transport = connect(pipe_name, connect_timeout_s=connect_timeout_s)
    except TransportError as exc:
        print(f"kernel: transport failed: {exc}", file=sys.stderr)
        return 2

    with transport:
        try:
            _handshake(transport)
        except (RuntimeError, FramingError, TransportError) as exc:
            print(f"kernel: handshake failed: {exc}", file=sys.stderr)
            return 3

        try:
            while True:
                frame = _recv(transport)
                if not _dispatch(transport, frame):
                    return 0
        except FramingError as exc:
            # Peer dropped or sent a malformed frame. Exit non-zero so the
            # supervisor knows this wasn't a clean shutdown.
            print(f"kernel: framing error: {exc}", file=sys.stderr)
            return 4
        except TransportError as exc:
            print(f"kernel: transport error: {exc}", file=sys.stderr)
            return 5


def main(argv: list[str] | None = None) -> NoReturn:
    """argv parser + entry. Always exits via :func:`sys.exit` — never returns."""
    import argparse

    parser = argparse.ArgumentParser(prog="pyexcel.kernel")
    parser.add_argument(
        "--pipe",
        required=True,
        help="server pipe name (without the platform-specific prefix)",
    )
    parser.add_argument(
        "--connect-timeout",
        type=float,
        default=5.0,
        help="seconds to wait for the supervisor's pipe to be ready",
    )
    args = parser.parse_args(argv)
    sys.exit(run(args.pipe, connect_timeout_s=args.connect_timeout))
