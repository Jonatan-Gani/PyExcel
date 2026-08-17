"""Tests for the worker's typed dispatch and named multi-output.

Covers the switch between the documented ``transform(inputs)`` dict
contract and the legacy positional one, the grid decode every declared type
is built from, and the dict return that previously reached Excel as the
literal string ``"StructArray"``.
"""

import os

import pandas as pd
import pytest

from pyexcel.kernel import arrow_io, worker


@pytest.fixture
def script(tmp_path):
    """Write a transform script and return its path."""

    def _write(body, name="s.py"):
        path = os.path.join(str(tmp_path), name)
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(body)
        return path

    return _write


def _meta(script_path, **extra):
    meta = {"run_id": "r", "script": script_path, "function": "transform"}
    meta.update(extra)
    return meta


# -----------------------------------------------------------------------------
# decode_grid — every shape arrives as a rectangle
# -----------------------------------------------------------------------------


def test_decode_grid_of_a_table():
    buf = arrow_io.encode(pd.DataFrame({"a": [1, 3], "b": [2, 4]}))
    assert arrow_io.decode_grid(buf) == [[1, 2], [3, 4]]


def test_decode_grid_of_a_column_vector():
    buf = arrow_io.encode([1, 2, 3], orientation=arrow_io.Orientation.COLUMN)
    assert arrow_io.decode_grid(buf) == [[1], [2], [3]]


def test_decode_grid_of_a_row_vector():
    buf = arrow_io.encode([1, 2, 3], orientation=arrow_io.Orientation.ROW)
    assert arrow_io.decode_grid(buf) == [[1, 2, 3]]


def test_decode_grid_of_a_scalar():
    assert arrow_io.decode_grid(arrow_io.encode(42)) == [[42]]


# -----------------------------------------------------------------------------
# Dispatch convention
# -----------------------------------------------------------------------------


def test_absent_inputs_key_keeps_positional_dispatch(script):
    """The UDF path and every pre-contract caller must be unaffected."""
    path = script("def transform(df):\n    return int(df.shape[0])\n")
    buf = arrow_io.encode(pd.DataFrame({"a": [1, 2]}))

    outcome = worker.run_job(_meta(path), [buf])

    assert outcome.success, outcome.meta
    assert arrow_io.decode(outcome.payloads[0]) == 2


def test_inputs_key_switches_to_the_dict_contract(script):
    path = script("def transform(inputs):\n    return sorted(inputs)\n")
    buf = arrow_io.encode(pd.DataFrame({"a": [1, 2]}))

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "Sales", "type": "dataframe"}]), [buf]
    )

    assert outcome.success, outcome.meta
    assert arrow_io.decode(outcome.payloads[0]) == ["Sales"]


def test_a_single_column_arrives_as_a_list_not_a_dataframe(script):
    """The regression that started this: A1:A10 of numbers is a list."""
    path = script(
        "def transform(inputs):\n"
        "    v = inputs['nums']\n"
        "    assert isinstance(v, list), type(v).__name__\n"
        "    return float(sum(v))\n"
    )
    buf = arrow_io.encode(pd.DataFrame({"0": [1, 2, 3, 4]}))

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "nums", "type": "auto"}]), [buf]
    )

    assert outcome.success, outcome.meta.get("message")
    assert arrow_io.decode(outcome.payloads[0]) == 10.0


@pytest.mark.parametrize(
    "value,expected_name",
    [(42, "value1"), ([1, 2, 3], "list1")],
)
def test_anonymous_bindings_are_auto_named_by_resolved_type(
    script, value, expected_name
):
    path = script("def transform(inputs):\n    return list(inputs)\n")
    buf = arrow_io.encode(value)

    outcome = worker.run_job(_meta(path, inputs=[{"type": "auto"}]), [buf])

    assert outcome.success, outcome.meta
    assert arrow_io.decode(outcome.payloads[0]) == [expected_name]


def test_payload_count_must_match_the_declared_bindings(script):
    path = script("def transform(inputs):\n    return 1\n")

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "a"}, {"name": "b"}]), [arrow_io.encode(1)]
    )

    assert not outcome.success
    assert "2 input binding" in outcome.meta["message"]


def test_input_contract_failure_names_the_binding(script):
    path = script("def transform(inputs):\n    return 1\n")
    # 2x2 — a real grid, which a Series genuinely cannot represent.
    buf = arrow_io.encode(pd.DataFrame({"a": [1, 3], "b": [2, 4]}))

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "Sales", "type": "series", "range": "A1:B2"}]),
        [buf],
    )

    assert not outcome.success
    assert outcome.meta["code"] == "BadInput"
    assert "Sales" in outcome.meta["message"]
    assert "A1:B2" in outcome.meta["message"]


# -----------------------------------------------------------------------------
# Named multi-output
# -----------------------------------------------------------------------------


def test_dict_return_becomes_one_named_payload_per_key(script):
    path = script(
        "def transform(inputs):\n"
        "    return {'total': 10.0, 'names': ['a', 'b']}\n"
    )

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "x", "type": "scalar"}]), [arrow_io.encode(1)]
    )

    assert outcome.success, outcome.meta
    assert outcome.meta["outputs"] == ["total", "names"]
    assert arrow_io.decode(outcome.payloads[0]) == 10.0
    assert arrow_io.decode(outcome.payloads[1]) == ["a", "b"]


def test_single_value_return_carries_no_output_names(script):
    path = script("def transform(inputs):\n    return 7\n")

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "x", "type": "scalar"}]), [arrow_io.encode(1)]
    )

    assert outcome.success
    assert "outputs" not in outcome.meta


def test_none_return_produces_no_payloads(script):
    path = script("def transform(inputs):\n    return None\n")

    outcome = worker.run_job(
        _meta(path, inputs=[{"name": "x", "type": "scalar"}]), [arrow_io.encode(1)]
    )

    assert outcome.success
    assert outcome.payloads == []


def test_declared_output_type_is_enforced(script):
    path = script("def transform(inputs):\n    return {'out': [1, 2, 3]}\n")

    outcome = worker.run_job(
        _meta(
            path,
            inputs=[{"name": "x", "type": "scalar"}],
            outputs=[{"name": "out", "type": "dataframe"}],
        ),
        [arrow_io.encode(1)],
    )

    assert not outcome.success
    assert outcome.meta["code"] == "BadReturnType"
    assert "out" in outcome.meta["message"]
    assert "DataFrame" in outcome.meta["message"]


def test_a_declared_output_missing_from_the_return_is_reported(script):
    path = script("def transform(inputs):\n    return {'other': 1}\n")

    outcome = worker.run_job(
        _meta(
            path,
            inputs=[{"name": "x", "type": "scalar"}],
            outputs=[{"name": "expected", "type": "scalar"}],
        ),
        [arrow_io.encode(1)],
    )

    assert not outcome.success
    assert "expected" in outcome.meta["message"]
    assert "other" in outcome.meta["message"]


def test_auto_typed_output_is_not_enforced(script):
    path = script("def transform(inputs):\n    return {'out': [1, 2, 3]}\n")

    outcome = worker.run_job(
        _meta(
            path,
            inputs=[{"name": "x", "type": "scalar"}],
            outputs=[{"name": "out", "type": "auto"}],
        ),
        [arrow_io.encode(1)],
    )

    assert outcome.success, outcome.meta
    assert arrow_io.decode(outcome.payloads[0]) == [1, 2, 3]


def test_malformed_binding_array_is_rejected(script):
    path = script("def transform(inputs):\n    return 1\n")

    outcome = worker.run_job(_meta(path, inputs="not-an-array"), [])

    assert not outcome.success
    assert outcome.meta["code"] == "BadRequest"
