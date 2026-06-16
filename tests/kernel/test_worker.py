"""Tests for ``pyexcel.kernel.worker``.

The worker is a pure function: meta + payloads in, JobOutcome out, no I/O.
That makes everything here a normal unit test — no subprocesses, no sockets.
End-to-end coverage (RUN_REQUEST over the wire driving the worker) lives
in ``test_supervisor.py``.
"""

from __future__ import annotations

import os
import textwrap
import time
from pathlib import Path

import pandas as pd
import pytest

from pyexcel.kernel import arrow_io, worker
from pyexcel.kernel.worker import JobOutcome, run_job


@pytest.fixture(autouse=True)
def _clear_module_cache():
    worker.clear_cache()
    yield
    worker.clear_cache()


@pytest.fixture(autouse=True)
def _reset_job_state():
    """Ensure the module-level per-job slots (cancel event, progress sink) are
    clear around every test, so a test that installs a sink can't leak it."""
    worker._end_job()
    yield
    worker._end_job()


def _write_script(tmp_path: Path, name: str, body: str) -> Path:
    p = tmp_path / name
    p.write_text(textwrap.dedent(body))
    return p


# -----------------------------------------------------------------------------
# Happy path: shape coverage
# -----------------------------------------------------------------------------


def test_scalar_in_scalar_out(tmp_path):
    script = _write_script(tmp_path, "addone.py", """
        def transform(x):
            return x + 1
    """)
    out = run_job(
        {"run_id": "r1", "script": str(script)},
        [arrow_io.encode(41)],
    )
    assert out.success
    assert out.meta["run_id"] == "r1"
    assert "duration_ms" in out.meta
    assert arrow_io.decode(out.payloads[0]) == 42


def test_dataframe_in_dataframe_out(tmp_path):
    script = _write_script(tmp_path, "doubler.py", """
        def transform(df):
            return df * 2
    """)
    df_in = pd.DataFrame({"x": [1, 2, 3]})
    out = run_job(
        {"run_id": "r2", "script": str(script)},
        [arrow_io.encode(df_in)],
    )
    assert out.success
    df_out = arrow_io.decode(out.payloads[0])
    assert isinstance(df_out, pd.DataFrame)
    assert df_out["x"].tolist() == [2, 4, 6]


def test_list_in_list_out(tmp_path):
    script = _write_script(tmp_path, "reverse.py", """
        def transform(values):
            return list(reversed(values))
    """)
    out = run_job(
        {"run_id": "r3", "script": str(script)},
        [arrow_io.encode([1, 2, 3])],
    )
    assert out.success
    assert arrow_io.decode(out.payloads[0]) == [3, 2, 1]


def test_no_args_returns_constant(tmp_path):
    script = _write_script(tmp_path, "const.py", """
        def transform():
            return 7
    """)
    out = run_job(
        {"run_id": "r4", "script": str(script)},
        [],
    )
    assert out.success
    assert arrow_io.decode(out.payloads[0]) == 7


def test_multiple_payloads_passed_as_positional_args(tmp_path):
    script = _write_script(tmp_path, "add.py", """
        def transform(a, b):
            return a + b
    """)
    out = run_job(
        {"run_id": "r5", "script": str(script)},
        [arrow_io.encode(10), arrow_io.encode(32)],
    )
    assert out.success
    assert arrow_io.decode(out.payloads[0]) == 42


def test_kwargs_from_meta_are_passed_through(tmp_path):
    script = _write_script(tmp_path, "kw.py", """
        def transform(x, *, factor):
            return x * factor
    """)
    out = run_job(
        {"run_id": "r6", "script": str(script), "kwargs": {"factor": 5}},
        [arrow_io.encode(8)],
    )
    assert out.success
    assert arrow_io.decode(out.payloads[0]) == 40


def test_none_return_yields_empty_payloads(tmp_path):
    script = _write_script(tmp_path, "noop.py", """
        def transform(x):
            return None
    """)
    out = run_job(
        {"run_id": "r7", "script": str(script)},
        [arrow_io.encode(1)],
    )
    assert out.success
    assert out.payloads == []


def test_custom_function_name_via_meta(tmp_path):
    script = _write_script(tmp_path, "custom.py", """
        def transform(x):
            return "wrong function"
        def my_func(x):
            return x.upper()
    """)
    out = run_job(
        {"run_id": "r8", "script": str(script), "function": "my_func"},
        [arrow_io.encode("hello")],
    )
    assert out.success
    assert arrow_io.decode(out.payloads[0]) == "HELLO"


# -----------------------------------------------------------------------------
# Meta contract
# -----------------------------------------------------------------------------


def test_run_id_echoed_in_success(tmp_path):
    script = _write_script(tmp_path, "id.py", "def transform(): return 0\n")
    out = run_job({"run_id": "my-unique-id", "script": str(script)}, [])
    assert out.meta["run_id"] == "my-unique-id"


def test_run_id_echoed_in_error(tmp_path):
    out = run_job({"run_id": "err-id", "script": "/no/such/file.py"}, [])
    assert not out.success
    assert out.meta["run_id"] == "err-id"


def test_duration_ms_is_non_negative_int(tmp_path):
    script = _write_script(tmp_path, "slow.py", """
        import time
        def transform():
            time.sleep(0.01)
            return None
    """)
    out = run_job({"run_id": "t", "script": str(script)}, [])
    assert isinstance(out.meta["duration_ms"], int)
    assert out.meta["duration_ms"] >= 0


# -----------------------------------------------------------------------------
# Error paths
# -----------------------------------------------------------------------------


def test_missing_script_meta_is_bad_request():
    out = run_job({"run_id": "x"}, [])
    assert not out.success
    assert out.meta["code"] == "BadRequest"
    assert "script" in out.meta["message"].lower()


def test_non_string_script_is_bad_request():
    out = run_job({"run_id": "x", "script": 123}, [])
    assert not out.success
    assert out.meta["code"] == "BadRequest"


def test_non_string_function_is_bad_request(tmp_path):
    script = _write_script(tmp_path, "s.py", "def transform(): return 0\n")
    out = run_job(
        {"run_id": "x", "script": str(script), "function": 42},
        [],
    )
    assert not out.success
    assert out.meta["code"] == "BadRequest"


def test_kwargs_not_dict_is_bad_request(tmp_path):
    script = _write_script(tmp_path, "s.py", "def transform(): return 0\n")
    out = run_job(
        {"run_id": "x", "script": str(script), "kwargs": [1, 2, 3]},
        [],
    )
    assert not out.success
    assert out.meta["code"] == "BadRequest"


def test_missing_script_file_is_module_not_found():
    out = run_job({"run_id": "x", "script": "/definitely/not/here.py"}, [])
    assert not out.success
    assert out.meta["code"] == "ModuleNotFound"


def test_script_with_import_error_is_module_exec_error(tmp_path):
    script = _write_script(tmp_path, "broken.py", """
        import this_module_definitely_does_not_exist
        def transform():
            return 1
    """)
    out = run_job({"run_id": "x", "script": str(script)}, [])
    assert not out.success
    assert out.meta["code"] == "ModuleExecError"
    assert out.meta["type"] == "ModuleNotFoundError"


def test_missing_function_is_function_not_found(tmp_path):
    script = _write_script(tmp_path, "no_transform.py", """
        def other():
            return 1
    """)
    out = run_job({"run_id": "x", "script": str(script)}, [])
    assert not out.success
    assert out.meta["code"] == "FunctionNotFound"


def test_non_callable_attribute_is_function_not_callable(tmp_path):
    script = _write_script(tmp_path, "not_callable.py", """
        transform = 42
    """)
    out = run_job({"run_id": "x", "script": str(script)}, [])
    assert not out.success
    assert out.meta["code"] == "FunctionNotCallable"


def test_bad_input_payload_is_bad_input(tmp_path):
    script = _write_script(tmp_path, "s.py", "def transform(x): return x\n")
    out = run_job(
        {"run_id": "x", "script": str(script)},
        [b"this is not arrow"],
    )
    assert not out.success
    assert out.meta["code"] == "BadInput"


def test_user_exception_is_surfaced_with_traceback(tmp_path):
    script = _write_script(tmp_path, "raises.py", """
        def transform(x):
            raise ValueError("boom")
    """)
    out = run_job(
        {"run_id": "x", "script": str(script)},
        [arrow_io.encode(1)],
    )
    assert not out.success
    assert out.meta["code"] == "Exception"
    assert out.meta["type"] == "ValueError"
    assert out.meta["message"] == "boom"
    assert "ValueError" in out.meta["traceback"]
    assert "raises.py" in out.meta["traceback"]


def test_unsupported_return_type_is_bad_return_type(tmp_path):
    script = _write_script(tmp_path, "weird.py", """
        class Opaque:
            pass
        def transform():
            return Opaque()
    """)
    out = run_job({"run_id": "x", "script": str(script)}, [])
    assert not out.success
    assert out.meta["code"] == "BadReturnType"


def test_missing_run_id_defaults_to_empty_string():
    out = run_job({"script": "/no/such/file.py"}, [])
    assert not out.success
    assert out.meta["run_id"] == ""


# -----------------------------------------------------------------------------
# Module caching
# -----------------------------------------------------------------------------


def test_unchanged_script_is_cached(tmp_path):
    # Probe by mutating a module-level counter inside the loaded module.
    # If the cache skips reload, the counter survives across calls.
    script = _write_script(tmp_path, "counter.py", """
        _calls = [0]
        def transform():
            _calls[0] += 1
            return _calls[0]
    """)
    a = run_job({"run_id": "1", "script": str(script)}, [])
    b = run_job({"run_id": "2", "script": str(script)}, [])
    assert arrow_io.decode(a.payloads[0]) == 1
    assert arrow_io.decode(b.payloads[0]) == 2  # state survived → module reused


def test_mtime_change_invalidates_cache(tmp_path):
    script = _write_script(tmp_path, "v.py", """
        VERSION = 1
        def transform():
            return VERSION
    """)
    out1 = run_job({"run_id": "1", "script": str(script)}, [])
    assert arrow_io.decode(out1.payloads[0]) == 1

    # Rewrite + bump mtime so the cache invalidates. Sleep briefly because
    # some filesystems have 1-second mtime resolution.
    time.sleep(1.1)
    script.write_text(textwrap.dedent("""
        VERSION = 2
        def transform():
            return VERSION
    """))

    out2 = run_job({"run_id": "2", "script": str(script)}, [])
    assert arrow_io.decode(out2.payloads[0]) == 2


# -----------------------------------------------------------------------------
# Determinism: failure result is still a well-typed JobOutcome
# -----------------------------------------------------------------------------


def test_outcome_is_frozen_dataclass():
    import dataclasses

    out = run_job({"run_id": "x"}, [])
    assert isinstance(out, JobOutcome)
    with pytest.raises(dataclasses.FrozenInstanceError):
        out.success = True  # type: ignore[misc]


# -----------------------------------------------------------------------------
# report_progress — user-facing helper; supervisor wires the sink, worker just
# forwards. (End-to-end PROGRESS-frame coverage lives in test_supervisor.py.)
# -----------------------------------------------------------------------------


def test_report_progress_is_noop_when_no_job_in_flight():
    # No _begin_job has installed a sink, so this must be inert (not raise).
    worker.report_progress(50, "ignored")
    worker.report_progress()  # bare call, indeterminate
    worker.report_progress(message="still fine")


def test_report_progress_forwards_to_installed_sink():
    seen: list[tuple] = []
    worker._begin_job(None, lambda pct, msg: seen.append((pct, msg)))
    try:
        worker.report_progress(25, "quarter")
        worker.report_progress(100, "done")
    finally:
        worker._end_job()

    assert seen == [(25.0, "quarter"), (100.0, "done")]


def test_report_progress_coerces_percent_to_float_and_message_to_str():
    seen: list[tuple] = []
    worker._begin_job(None, lambda pct, msg: seen.append((pct, msg)))
    try:
        worker.report_progress(42, 123)  # int percent, non-str message
    finally:
        worker._end_job()

    (pct, msg), = seen
    assert isinstance(pct, float) and pct == 42.0
    assert msg == "123"


def test_report_progress_none_percent_is_indeterminate():
    seen: list[tuple] = []
    worker._begin_job(None, lambda pct, msg: seen.append((pct, msg)))
    try:
        worker.report_progress(message="working")  # percent defaults to None
        worker.report_progress(None, "still working")
    finally:
        worker._end_job()

    assert seen == [(None, "working"), (None, "still working")]


def test_end_job_clears_sink_so_later_calls_are_noops():
    seen: list[tuple] = []
    worker._begin_job(None, lambda pct, msg: seen.append((pct, msg)))
    worker.report_progress(10, "during")
    worker._end_job()
    worker.report_progress(90, "after")  # sink gone — must not be recorded

    assert seen == [(10.0, "during")]


# -----------------------------------------------------------------------------
# Standard-input guard
# -----------------------------------------------------------------------------


@pytest.fixture
def _input_guard():
    """Install the input guard for one test and restore process globals after,
    so the session-wide ``builtins.input`` / ``sys.stdin`` (and pytest's own
    capture) aren't left mutated."""
    import builtins
    import sys

    saved_input, saved_stdin = builtins.input, sys.stdin
    worker._input_guard_installed = False
    worker.install_input_guard()
    try:
        yield
    finally:
        builtins.input = saved_input
        sys.stdin = saved_stdin
        worker._input_guard_installed = False


def test_input_call_fails_fast_with_explanation(tmp_path, _input_guard):
    script = _write_script(tmp_path, "asks.py", """
        def transform(x):
            return input()
    """)
    out = run_job({"run_id": "r1", "script": str(script)}, [arrow_io.encode(1)])

    assert not out.success
    assert out.meta["type"] == "PyExcelInputError"
    assert "input() is disabled" in out.meta["message"]
    assert out.meta["run_id"] == "r1"


def test_stdin_read_fails_fast(tmp_path, _input_guard):
    script = _write_script(tmp_path, "reads.py", """
        import sys

        def transform(x):
            return sys.stdin.readline()
    """)
    out = run_job({"run_id": "r2", "script": str(script)}, [arrow_io.encode(1)])

    assert not out.success
    assert out.meta["type"] == "PyExcelInputError"


def test_install_input_guard_is_idempotent(_input_guard):
    # A second install must not raise or re-wrap; input() stays blocked.
    worker.install_input_guard()
    with pytest.raises(worker.PyExcelInputError):
        input()


def test_normal_script_unaffected_by_guard(tmp_path, _input_guard):
    # The guard only touches stdin — ordinary transforms still run.
    script = _write_script(tmp_path, "addone.py", """
        def transform(x):
            return x + 1
    """)
    out = run_job({"run_id": "r3", "script": str(script)}, [arrow_io.encode(41)])

    assert out.success
    assert arrow_io.decode(out.payloads[0]) == 42
