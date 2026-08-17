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

import builtins
import dataclasses
import hashlib
import importlib.util
import io
import os
import sys
import threading
import time
import traceback
from types import ModuleType
from typing import Any, Callable, Dict, List, Optional, Tuple

from . import arrow_io
from . import declared_types


_DEFAULT_FUNCTION = "transform"


# -----------------------------------------------------------------------------
# Standard-input guard
#
# The kernel is spawned by the Excel host as a subprocess with no interactive
# console attached (the host does not redirect stdin, so the child inherits a
# handle that never delivers a line). A user script that calls ``input()`` — or
# otherwise reads ``sys.stdin`` — therefore blocks the worker thread forever:
# the supervisor's run loop never sees the job finish, never writes a reply, and
# the host eventually fails the run with an opaque "no frame received" timeout.
#
# :func:`install_input_guard` turns that hang into an immediate, explained error
# by replacing ``builtins.input`` and ``sys.stdin`` before any user code runs.
# Nothing in the kernel legitimately reads stdin (the wire is a named pipe), so
# disabling it process-wide is safe.
# -----------------------------------------------------------------------------

_INPUT_FORBIDDEN_MESSAGE = (
    "input() is disabled in PyExcel scripts. The kernel runs without an "
    "interactive console, so input() (or any read from standard input) would "
    "block the run until it times out. Remove the input() call — read your "
    "values from the transform() 'inputs' argument instead."
)


class PyExcelInputError(RuntimeError):
    """Raised when user code calls ``input()`` or reads ``sys.stdin`` inside
    the kernel. Surfaced to the host as the ``type`` of a clean ERROR frame so
    the user sees the explanation in :data:`_INPUT_FORBIDDEN_MESSAGE` rather
    than a 60-second timeout. See :func:`install_input_guard`.
    """


class _ForbiddenStdin(io.IOBase):
    """Stand-in for ``sys.stdin`` that fails fast instead of blocking.

    Every read raises :class:`PyExcelInputError`; ``isatty()`` reports False so
    libraries probing for an interactive terminal behave as in a non-interactive
    environment instead of trying to read.
    """

    def readable(self) -> bool:
        return False

    def isatty(self) -> bool:
        return False

    def _blocked(self, *args: Any, **kwargs: Any) -> Any:
        raise PyExcelInputError(_INPUT_FORBIDDEN_MESSAGE)

    read = _blocked
    readline = _blocked
    readlines = _blocked


_input_guard_installed = False


def install_input_guard() -> None:
    """Make ``input()`` and ``sys.stdin`` reads fail fast in the kernel process.

    Called once by the supervisor before the run loop starts. Idempotent — safe
    to call more than once. After this, a user ``input()`` call raises
    :class:`PyExcelInputError` immediately, which :func:`run_job` folds into a
    clean ERROR outcome the host can show.
    """
    global _input_guard_installed
    if _input_guard_installed:
        return

    def _forbidden_input(prompt: object = "", /) -> str:  # noqa: ARG001
        raise PyExcelInputError(_INPUT_FORBIDDEN_MESSAGE)

    builtins.input = _forbidden_input  # type: ignore[assignment]
    try:
        sys.stdin = _ForbiddenStdin()  # type: ignore[assignment]
    except Exception:  # pragma: no cover - extremely defensive
        # Reassigning sys.stdin should never fail, but if it somehow does the
        # builtins.input override above is the one that matters most.
        pass
    _input_guard_installed = True

# abs_path -> (mtime_seen_at_load, loaded_module)
_module_cache: dict[str, Tuple[float, ModuleType]] = {}

# Per-job state, kept thread-local so it is scoped to the worker thread that
# runs the job. Only one job runs at a time in the common case, but when a
# cancelled, non-cooperative worker is abandoned (see
# :func:`pyexcel.kernel.supervisor._run_with_cancellation`) the abandoned thread
# keeps its own cancel-event/progress-sink — so when it eventually unblocks and
# finishes it can't disturb the next job, which runs on a fresh worker thread
# with its own slot. ``cancel_event`` is set when a CANCEL frame arrives mid-run
# (:func:`is_cancelled` reads it); ``progress_sink`` is the callback
# :func:`report_progress` forwards user progress updates to.
ProgressSink = Callable[[Optional[float], str], None]
_job_state = threading.local()


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
    ev = getattr(_job_state, "cancel_event", None)
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
    sink = getattr(_job_state, "progress_sink", None)
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
    _job_state.cancel_event = event
    _job_state.progress_sink = progress_sink


def _end_job() -> None:
    """Clear the per-job cancellation Event + progress sink after the worker
    thread finishes, so a later out-of-band :func:`report_progress` /
    :func:`is_cancelled` call is an inert no-op."""
    _job_state.cancel_event = None
    _job_state.progress_sink = None


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

        # The presence of an 'inputs' array is the switch between the typed
        # dict contract and the legacy positional one. The UDF path and
        # direct kernel-client callers send no 'inputs' and keep working.
        input_specs = _input_specs(req_meta)
        if input_specs is None:
            args = _decode_args(req_payloads)
            value = fn(*args, **kwargs)
        else:
            value = fn(_build_inputs(input_specs, req_payloads), **kwargs)

        payloads, output_names = _encode_outputs(value, _output_specs(req_meta))

        meta = {
            "run_id": run_id,
            "duration_ms": _elapsed_ms(started, _now),
        }
        if output_names is not None:
            meta["outputs"] = output_names

        return JobOutcome(success=True, meta=meta, payloads=payloads)
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
# Typed I/O contract — see docs/typed-io-contract.md
# -----------------------------------------------------------------------------

#: Prefix used to auto-name an anonymous binding, keyed by resolved type.
#: Matches what README.md advertises (df1, list1, value1, ...).
_AUTO_NAME_PREFIX = {
    declared_types.DATAFRAME: "df",
    declared_types.SERIES: "series",
    declared_types.LIST: "list",
    declared_types.TUPLE: "list",
    declared_types.SET: "list",
    declared_types.NDARRAY: "list",
    declared_types.DICT: "dict",
    declared_types.SCALAR: "value",
}


def _binding_specs(req_meta: dict, key: str) -> Optional[List[dict]]:
    """Read an 'inputs'/'outputs' binding array off the request meta.

    Returns ``None`` when the key is absent — the signal that this caller
    predates the typed contract and wants the legacy behaviour.
    """
    raw = req_meta.get(key)
    if raw is None:
        return None
    if not isinstance(raw, list):
        raise JobError("BadRequest", f"'{key}' must be a JSON array")

    specs: List[dict] = []
    for i, entry in enumerate(raw):
        if not isinstance(entry, dict):
            raise JobError("BadRequest", f"'{key}[{i}]' must be a JSON object")
        specs.append(entry)
    return specs


def _input_specs(req_meta: dict) -> Optional[List[dict]]:
    return _binding_specs(req_meta, "inputs")


def _output_specs(req_meta: dict) -> Optional[List[dict]]:
    return _binding_specs(req_meta, "outputs")


def _resolved_name(spec: dict, declared: str, counters: dict) -> str:
    """The binding's name, auto-generated from its type when unnamed."""
    name = spec.get("name")
    if isinstance(name, str) and name.strip():
        return name.strip()

    prefix = _AUTO_NAME_PREFIX.get(declared, "value")
    counters[prefix] = counters.get(prefix, 0) + 1
    return f"{prefix}{counters[prefix]}"


def _build_inputs(specs: List[dict], payloads: List[bytes]) -> Dict[str, Any]:
    """Decode each payload to a grid and construct its declared type.

    The result is the ``inputs`` dict the documented ``transform(inputs)``
    signature receives.
    """
    if len(specs) != len(payloads):
        raise JobError(
            "BadRequest",
            f"request declares {len(specs)} input binding(s) but carries "
            f"{len(payloads)} payload(s)",
        )

    inputs: Dict[str, Any] = {}
    counters: dict = {}

    for i, (spec, payload) in enumerate(zip(specs, payloads)):
        declared = spec.get("type") or declared_types.AUTO
        address = spec.get("range") or ""

        try:
            grid = arrow_io.decode_grid(payload)
        except BaseException as exc:  # noqa: BLE001 — pyarrow errors aren't a fixed type
            raise JobError(
                "BadInput",
                f"could not decode input {i} ('{spec.get('name') or address}'): {exc}",
                cause=exc,
            ) from exc

        # Resolve auto against the real grid before naming, so an anonymous
        # single-cell binding is named value1 rather than df1.
        rows = len(grid)
        columns = len(grid[0]) if rows else 0
        if declared == declared_types.AUTO:
            declared = declared_types.resolve_auto(rows, columns)

        name = _resolved_name(spec, declared, counters)

        try:
            inputs[name] = declared_types.build(
                declared, grid, name=name, address=address
            )
        except declared_types.TypeContractError as exc:
            raise JobError("BadInput", str(exc), cause=exc) from exc

    return inputs


def _encode_outputs(
    value: Any, specs: Optional[List[dict]]
) -> Tuple[List[bytes], Optional[List[str]]]:
    """Encode a return value into payloads, honouring declared output types.

    Returns ``(payloads, names)``. ``names`` is ``None`` for the single
    anonymous result the legacy path produces, and a list parallel to
    ``payloads`` when the return is a dict of named results — which is what
    lets the host route each one to its own output range.
    """
    if value is None:
        return [], None

    if isinstance(value, dict):
        return _encode_dict_outputs(value, specs)

    # A single value. Enforce the first declared output type, if any.
    if specs:
        declared = specs[0].get("type")
        name = specs[0].get("name") or "result"
        try:
            declared_types.check_output(declared, value, name=name)
        except declared_types.TypeContractError as exc:
            raise JobError("BadReturnType", str(exc), cause=exc) from exc

    return _encode_result(value), None


def _encode_dict_outputs(
    value: dict, specs: Optional[List[dict]]
) -> Tuple[List[bytes], List[str]]:
    """Encode a dict return as one named payload per key.

    Before this contract a returned dict fell through arrow_io's catch-all
    and reached Excel as the literal string "StructArray" — while both
    scaffolded templates and README told users to return exactly that.
    """
    declared_by_name = {}
    if specs:
        for spec in specs:
            spec_name = spec.get("name")
            if isinstance(spec_name, str) and spec_name.strip():
                declared_by_name[spec_name.strip()] = spec.get("type")

        missing = [n for n in declared_by_name if n not in value]
        if missing:
            raise JobError(
                "BadReturnType",
                "transform() returned no value for declared output "
                + ", ".join(f"'{n}'" for n in missing)
                + ". Returned keys: "
                + (", ".join(f"'{k}'" for k in value) or "(none)"),
            )

    payloads: List[bytes] = []
    names: List[str] = []

    for key, item in value.items():
        name = str(key)
        try:
            declared_types.check_output(declared_by_name.get(name), item, name=name)
        except declared_types.TypeContractError as exc:
            raise JobError("BadReturnType", str(exc), cause=exc) from exc

        try:
            payloads.append(arrow_io.encode(item))
        except TypeError as exc:
            raise JobError(
                "BadReturnType",
                f"could not encode output '{name}' of type "
                f"{type(item).__name__}: {exc}",
                cause=exc,
            ) from exc
        names.append(name)

    return payloads, names


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
