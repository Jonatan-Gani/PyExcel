"""
Kernel-side job execution.

The supervisor reads a ``RUN_REQUEST`` frame off the wire and hands it to
:func:`run_job`, which is a pure function: given a request meta + payloads,
it returns a :class:`JobOutcome` describing the reply frame the supervisor
should send back. No I/O happens here.

Wire contract (request meta):

    {
        "run_id":   str,            # required; echoed back in every reply
        "script":   str,            # required; absolute or cwd-relative .py
        "function": str = "transform",
        "kwargs":   dict = {},      # JSON-serialisable
    }

Each payload in the request is an Arrow IPC stream produced by
:mod:`pyexcel.kernel.arrow_io`. Payloads decode to positional arguments
(``fn(*args, **kwargs)``) in order.

Wire contract (success reply):

    type     = RUN_RESULT
    meta     = {"run_id": str, "duration_ms": int}
    payloads = []           if the user function returned ``None``
               [arrow_buf]  otherwise (single Arrow IPC stream)

Wire contract (failure reply):

    type     = ERROR
    meta     = {
        "run_id":      str,         # echoed; empty if the request lacked it
        "code":        str,         # see "Error codes" below
        "type":        str,         # Python exception class name
        "message":     str,         # exception message
        "traceback":   str,         # formatted traceback
        "duration_ms": int,         # wall time spent before the failure
    }
    payloads = []

Error codes:

    BadRequest       request meta is missing or malformed
    ModuleNotFound   script path does not exist
    ModuleExecError  script raised during import
    FunctionNotFound function attribute missing on module
    FunctionNotCallable function attribute exists but isn't callable
    BadInput         a request payload failed to Arrow-decode
    BadReturnType    user return value can't be Arrow-encoded
    Exception        anything else the user function raised

Module loading caches by absolute path + mtime so repeated calls against
the same unchanged script don't re-exec. Touching the file invalidates
the cache on the next call.
"""

from __future__ import annotations

import dataclasses
import hashlib
import importlib.util
import os
import sys
import threading
import time
import traceback
from types import ModuleType
from typing import Any, Callable, List, Optional, Tuple

from . import arrow_io


_DEFAULT_FUNCTION = "transform"

# abs_path -> (mtime_seen_at_load, loaded_module)
_module_cache: dict[str, Tuple[float, ModuleType]] = {}

# Module-level per-job state. Only one job runs at a time per kernel, so a
# single shared slot for each is sufficient — the supervisor calls
# :func:`_begin_job` before dispatching and :func:`_end_job` afterwards.
#
# ``_current_cancel_event`` is set when a CANCEL frame arrives mid-run;
# :func:`is_cancelled` reads it. ``_current_progress_sink`` is the callback
# :func:`report_progress` forwards user progress updates to — the supervisor
# wires it to a queue it drains onto the wire as PROGRESS frames.
_current_cancel_event: Optional[threading.Event] = None
ProgressSink = Callable[[Optional[float], str], None]
_current_progress_sink: Optional[ProgressSink] = None


class JobError(Exception):
    """Internal exception type used to short-circuit ``run_job`` with a
    well-defined error code. Never escapes — :func:`run_job` always catches
    it and returns a failure :class:`JobOutcome`.
    """

    def __init__(self, code: str, message: str = "", *, cause: BaseException | None = None) -> None:
        super().__init__(message or code)
        self.code = code
        self.cause = cause


@dataclasses.dataclass(frozen=True)
class JobOutcome:
    """The reply the supervisor should write back to the host.

    ``success`` selects the frame type (``RUN_RESULT`` vs ``ERROR``);
    ``meta`` and ``payloads`` go straight into the frame.
    """

    success: bool
    meta: dict
    payloads: List[bytes]


def clear_cache() -> None:
    """Drop every cached user module. Test helper; not used at runtime."""
    _module_cache.clear()


def is_cancelled() -> bool:
    """User-facing API: was a CANCEL frame received for the current job?

    Long-running transform functions can poll this between work units and
    return early — the supervisor will then surface an ``ERROR`` frame with
    code ``"Cancelled"`` instead of ``RUN_RESULT``. Returns ``False`` when
    no job is in flight, so it's always safe to call.
    """
    ev = _current_cancel_event
    return ev is not None and ev.is_set()


def report_progress(percent: Optional[float] = None, message: str = "") -> None:
    """User-facing API: report progress for the in-flight job to the host.

    Long-running transform functions can call this between work units; the
    supervisor relays each call to the host as a ``PROGRESS`` frame
    (``KernelClient.ProgressReceived`` on the C# side), which a progress UI
    renders. Calls are fire-and-forget — they never block on the host and
    never raise on the user for transport reasons.

    Args:
        percent: Completion as a 0–100 value, or ``None`` for an
            indeterminate update that only carries a ``message`` (e.g. a
            spinner with a status line). Numeric values are coerced to
            ``float``; the 0–100 convention is honoured by the renderer, not
            enforced here.
        message: Optional human-readable status line.

    Calling this when no job is in flight is a safe no-op (mirrors
    :func:`is_cancelled`), so user modules can call it unconditionally.
    """
    sink = _current_progress_sink
    if sink is None:
        return  # no job in flight — safe no-op
    pct = None if percent is None else float(percent)
    sink(pct, str(message))


def _begin_job(
    event: Optional[threading.Event],
    progress_sink: Optional[ProgressSink] = None,
) -> None:
    """Install the per-job cancellation Event + progress sink for one job.

    Called by :mod:`pyexcel.kernel.supervisor` immediately before dispatching
    ``run_job`` on the worker thread. Not part of the user-facing surface.
    """
    global _current_cancel_event, _current_progress_sink
    _current_cancel_event = event
    _current_progress_sink = progress_sink


def _end_job() -> None:
    """Clear the per-job cancellation Event + progress sink after the worker
    thread finishes, so a later out-of-band :func:`report_progress` /
    :func:`is_cancelled` call is an inert no-op."""
    global _current_cancel_event, _current_progress_sink
    _current_cancel_event = None
    _current_progress_sink = None


def run_job(
    req_meta: dict,
    req_payloads: List[bytes],
    *,
    _now: Callable[[], float] = time.monotonic,
) -> JobOutcome:
    """Execute one job. Never raises — every failure is folded into the
    returned :class:`JobOutcome`.

    ``_now`` is injectable for deterministic duration assertions in tests.
    """
    run_id = req_meta.get("run_id") or ""
    started = _now()

    try:
        script = req_meta.get("script")
        if not script or not isinstance(script, str):
            raise JobError("BadRequest", "request meta is missing or has non-string 'script'")

        function_name = req_meta.get("function") or _DEFAULT_FUNCTION
        if not isinstance(function_name, str):
            raise JobError("BadRequest", "'function' must be a string")

        kwargs = req_meta.get("kwargs") or {}
        if not isinstance(kwargs, dict):
            raise JobError("BadRequest", "'kwargs' must be a JSON object")

        module = _load_module(script)
        fn = _resolve_function(module, function_name)
        args = _decode_args(req_payloads)

        value = fn(*args, **kwargs)
        payloads = _encode_result(value)

        return JobOutcome(
            success=True,
            meta={
                "run_id": run_id,
                "duration_ms": _elapsed_ms(started, _now),
            },
            payloads=payloads,
        )
    except JobError as exc:
        return JobOutcome(
            success=False,
            meta=_error_meta(run_id, exc.code, exc.cause or exc, _elapsed_ms(started, _now)),
            payloads=[],
        )
    except BaseException as exc:  # noqa: BLE001 — last-resort catch
        # User code blew up. Surface it as a clean ERROR rather than letting
        # the supervisor's frame loop die.
        return JobOutcome(
            success=False,
            meta=_error_meta(run_id, "Exception", exc, _elapsed_ms(started, _now)),
            payloads=[],
        )


# -----------------------------------------------------------------------------
# Module loading
# -----------------------------------------------------------------------------


def _module_name_for(abs_path: str) -> str:
    """Deterministic-within-process module name keyed by absolute path.

    SHA-1-truncated rather than ``hash()`` because the latter changes with
    ``PYTHONHASHSEED`` and can theoretically collide more often.
    """
    digest = hashlib.sha1(abs_path.encode("utf-8")).hexdigest()[:16]
    return f"_pyexcel_user_{digest}"


def _load_module(script_path: str) -> ModuleType:
    abs_path = os.path.abspath(script_path)
    if not os.path.isfile(abs_path):
        raise JobError("ModuleNotFound", f"no such script: {abs_path}")

    mtime = os.path.getmtime(abs_path)
    cached = _module_cache.get(abs_path)
    if cached is not None and cached[0] == mtime:
        return cached[1]

    mod_name = _module_name_for(abs_path)
    spec = importlib.util.spec_from_file_location(mod_name, abs_path)
    if spec is None or spec.loader is None:
        raise JobError("ModuleLoadError", f"could not create import spec for {abs_path}")

    module = importlib.util.module_from_spec(spec)
    sys.modules[mod_name] = module
    try:
        spec.loader.exec_module(module)
    except BaseException as exc:  # noqa: BLE001 — preserve the original cause
        sys.modules.pop(mod_name, None)
        raise JobError(
            "ModuleExecError",
            f"failed to load {abs_path}",
            cause=exc,
        ) from exc

    _module_cache[abs_path] = (mtime, module)
    return module


def _resolve_function(module: ModuleType, function_name: str) -> Any:
    fn = getattr(module, function_name, None)
    if fn is None:
        raise JobError(
            "FunctionNotFound",
            f"function {function_name!r} not found in {module.__name__}",
        )
    if not callable(fn):
        raise JobError(
            "FunctionNotCallable",
            f"{function_name!r} on {module.__name__} is not callable",
        )
    return fn


# -----------------------------------------------------------------------------
# Payload (de)serialisation
# -----------------------------------------------------------------------------


def _decode_args(payloads: List[bytes]) -> List[Any]:
    args: List[Any] = []
    for i, p in enumerate(payloads):
        try:
            args.append(arrow_io.decode(p))
        except BaseException as exc:  # noqa: BLE001 — pyarrow errors aren't a fixed type
            raise JobError(
                "BadInput",
                f"could not decode payload {i}: {exc}",
                cause=exc,
            ) from exc
    return args


def _encode_result(value: Any) -> List[bytes]:
    if value is None:
        return []
    try:
        return [arrow_io.encode(value)]
    except TypeError as exc:
        raise JobError(
            "BadReturnType",
            f"could not encode return value of type {type(value).__name__}: {exc}",
            cause=exc,
        ) from exc


# -----------------------------------------------------------------------------
# Meta helpers
# -----------------------------------------------------------------------------


def _elapsed_ms(started: float, now: Callable[[], float]) -> int:
    return int(round((now() - started) * 1000))


def _error_meta(run_id: str, code: str, exc: BaseException, duration_ms: int) -> dict:
    return {
        "run_id": run_id,
        "code": code,
        "type": type(exc).__name__,
        "message": str(exc),
        "traceback": "".join(
            traceback.format_exception(type(exc), exc, exc.__traceback__)
        ),
        "duration_ms": duration_ms,
    }
