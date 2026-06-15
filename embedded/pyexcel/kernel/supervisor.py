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
* ``RUN_REQUEST`` — delegated to :mod:`pyexcel.kernel.worker` running on
  a worker thread; meanwhile the main loop pumps inbound frames so
  ``CANCEL`` arrives in time to flip the cooperative cancellation flag,
  and ``PING`` keeps answering for liveness checks. The same pump drains
  any ``pyexcel.kernel.report_progress`` calls the user function made and
  emits them as ``PROGRESS`` frames — all wire writes stay on the main
  thread, so the worker thread never races the loop for the transport.
  Replies inline with ``RUN_RESULT`` or ``ERROR`` once the worker thread
  finishes.
* ``SHUTDOWN`` — clean exit with code 0.

Anything else gets an ``ERROR`` reply and the loop stays alive.
"""

from __future__ import annotations

import queue
import sys
import threading
import time
from typing import NoReturn, Optional

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

# How often the main loop wakes to poll for inbound frames while a
# RUN_REQUEST is executing on the worker thread. 50 ms gives sub-100ms
# CANCEL latency with near-idle CPU. Tunable per-test by patching.
_CANCEL_POLL_INTERVAL_S = 0.05

# How long the main loop waits for the worker thread to wind down once
# the worker function has either returned naturally or noticed a CANCEL
# request. The worker thread should exit promptly after run_job returns;
# this guard rail exists to keep an unresponsive script from wedging the
# supervisor.
_WORKER_JOIN_TIMEOUT_S = 5.0

# After a CANCEL arrives, how long to let the worker wind down cooperatively
# before giving up on it. A cooperative script that polls
# :func:`pyexcel.kernel.is_cancelled` returns well within this; a
# non-cooperative one (blocked in sleep/IO, or ignoring the flag) is abandoned
# once it lapses so the host gets a prompt ERROR/Cancelled instead of waiting
# out its full run deadline. Kept short so Cancel feels responsive.
_CANCEL_GRACE_S = 0.5


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
        _run_with_cancellation(transport, frame)
        return True

    if frame.type is FrameType.SHUTDOWN:
        return False

    if frame.type is FrameType.CANCEL:
        # An unsolicited CANCEL (no run in flight) is a no-op — the host
        # may have raced us and sent it after the previous run already
        # finished. Acknowledge with ERROR so the host knows it was seen
        # but didn't apply; the loop stays alive.
        _send(
            transport,
            FrameType.ERROR,
            {
                "reason": "CANCEL received but no run in flight",
                "code": "Cancelled",
                "run_id": frame.meta.get("run_id", ""),
            },
        )
        return True

    # Anything not in the dispatch table above (LIST_JOBS, etc.) gets a
    # polite ERROR rather than silent drop. The C# side surfaces this in tests.
    _send(
        transport,
        FrameType.ERROR,
        {"reason": f"unsupported frame type {frame.type.name} at this phase"},
    )
    return True


def _flush_progress(
    transport: FrameTransport,
    progress_queue: "queue.Queue[tuple[Optional[float], str]]",
    run_id: str,
) -> None:
    """Drain every queued ``report_progress`` call onto the wire as PROGRESS
    frames. Called from the main loop only, so it never races the worker
    thread for the transport. Meta mirrors what ``KernelClient.RaiseProgress``
    reads: ``run_id`` echoed, ``percent`` (``None`` for indeterminate),
    ``message``.
    """
    while True:
        try:
            percent, message = progress_queue.get_nowait()
        except queue.Empty:
            return
        _send(
            transport,
            FrameType.PROGRESS,
            {"run_id": run_id, "percent": percent, "message": message},
        )


def _run_with_cancellation(transport: FrameTransport, request: Frame) -> None:
    """Execute one RUN_REQUEST on a worker thread, pumping inbound frames in
    the main loop so CANCEL arrives in time to flip the cooperative
    cancellation flag and PING keeps answering for liveness checks. The pump
    also drains the worker's ``report_progress`` calls and emits them as
    PROGRESS frames, so every wire write stays on this (the main) thread.

    Reply is sent once after the worker thread finishes. If a CANCEL arrived
    during the run, the reply is overridden to ``ERROR / Cancelled`` regardless
    of whether the user function completed naturally or noticed and aborted.
    """
    cancel_event = threading.Event()
    # Bounded by nothing in principle, but the pump drains it every poll tick;
    # the worker only ever enqueues from its own thread. Tuple is (percent, msg).
    progress_queue: "queue.Queue[tuple[Optional[float], str]]" = queue.Queue()
    outcome_holder: list[Optional[worker.JobOutcome]] = [None]
    run_id = request.meta.get("run_id", "") or ""
    started = time.monotonic()

    def progress_sink(percent: Optional[float], message: str) -> None:
        # Runs on the worker thread; only enqueues. The main loop does the
        # actual transport write so we keep a single writer.
        progress_queue.put((percent, message))

    def worker_main() -> None:
        worker._begin_job(cancel_event, progress_sink)
        try:
            outcome_holder[0] = worker.run_job(request.meta, request.payloads)
        finally:
            worker._end_job()

    t = threading.Thread(target=worker_main, daemon=True, name="pyexcel-worker")
    t.start()

    # Pump inbound frames while the worker runs. CANCEL → flip the flag and
    # start a short grace timer; PING → reply with PONG; anything else
    # (including a second RUN_REQUEST, which the host shouldn't send while a run
    # is in flight) → ignore with a log so we don't deadlock waiting for a frame
    # we can't service yet. Each tick also flushes queued progress so updates
    # stream during the run rather than bunching up at the end.
    #
    # Once cancelled, we wait only until ``cancel_deadline`` for the worker to
    # wind down. A cooperative worker returns before then (and the loop exits on
    # ``t.is_alive()``); a non-cooperative one (blocked in sleep/IO, or ignoring
    # the flag) would otherwise pin this loop until the host's run deadline, so
    # we stop waiting and abandon it below.
    cancel_deadline: Optional[float] = None
    while t.is_alive():
        _flush_progress(transport, progress_queue, run_id)
        if cancel_deadline is not None and time.monotonic() >= cancel_deadline:
            break
        if not transport.has_data(_CANCEL_POLL_INTERVAL_S):
            continue
        try:
            f = _recv(transport)
        except (FramingError, TransportError) as exc:
            # Peer dropped mid-run. Set the cancel flag so the worker can
            # bail at its next checkpoint, and stop pumping; the outer loop
            # in :func:`run` will surface the disconnect on its next read.
            print(f"kernel: peer dropped during run {run_id!r}: {exc}", file=sys.stderr)
            cancel_event.set()
            break

        if f.type is FrameType.CANCEL:
            cancel_event.set()
            if cancel_deadline is None:
                cancel_deadline = time.monotonic() + _CANCEL_GRACE_S
        elif f.type is FrameType.PING:
            _send(transport, FrameType.PONG, {"nonce": f.meta.get("nonce", "")})
        else:
            print(
                f"kernel: ignoring {f.type.name} frame during run {run_id!r}",
                file=sys.stderr,
            )

    # If the worker is still running, we left the pump because a cancel grace
    # lapsed (or the peer dropped). We can't safely interrupt a Python thread,
    # so abandon it: it's a daemon and its job state is thread-local, so a late
    # finish can't disturb the next run. Reply now rather than blocking on a
    # join that may never complete and leaving the host to time out.
    if t.is_alive():
        _flush_progress(transport, progress_queue, run_id)
        _send(
            transport,
            FrameType.ERROR,
            {
                "run_id": run_id,
                "code": "Cancelled",
                "type": "CancellationRequested",
                "message": (
                    "run cancelled; the worker did not stop within "
                    f"{_CANCEL_GRACE_S}s and was abandoned"
                ),
                "traceback": "",
                "duration_ms": int(round((time.monotonic() - started) * 1000)),
            },
        )
        return

    t.join(timeout=_WORKER_JOIN_TIMEOUT_S)
    # Flush progress enqueued in the worker's final stretch before the loop
    # last polled, so every PROGRESS frame precedes the terminal reply.
    _flush_progress(transport, progress_queue, run_id)
    outcome = outcome_holder[0]

    # Three terminal states for a worker that finished:
    #   (a) returned normally + no cancel → forward the outcome as-is.
    #   (b) returned normally + cancel flag set → override with Cancelled.
    #   (c) did not return within the join timeout → WorkerHung error.
    if outcome is None:
        reply_meta = {
            "run_id": run_id,
            "code": "WorkerHung",
            "type": "TimeoutError",
            "message": f"worker thread did not finish within {_WORKER_JOIN_TIMEOUT_S}s",
            "traceback": "",
            "duration_ms": int(_WORKER_JOIN_TIMEOUT_S * 1000),
        }
        _send(transport, FrameType.ERROR, reply_meta)
        return

    if cancel_event.is_set():
        reply_meta = {
            "run_id": run_id,
            "code": "Cancelled",
            "type": "CancellationRequested",
            "message": "kernel received CANCEL during run",
            "traceback": "",
            "duration_ms": outcome.meta.get("duration_ms", 0),
        }
        _send(transport, FrameType.ERROR, reply_meta)
        return

    reply_type = FrameType.RUN_RESULT if outcome.success else FrameType.ERROR
    _send(transport, reply_type, outcome.meta, tuple(outcome.payloads))


def run(pipe_name: str, *, connect_timeout_s: float = 5.0) -> int:
    """Connect, handshake, and run the main loop. Returns the process exit code."""
    # Disable stdin before any user code can run: the kernel has no console, so
    # input() / sys.stdin reads would block the worker thread forever and the
    # host would only see a "no frame received" timeout. With the guard they
    # fail fast with a clear, surfaced error instead.
    worker.install_input_guard()

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
