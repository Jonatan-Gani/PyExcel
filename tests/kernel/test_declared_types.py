"""Tests for the declared-type coercion matrix.

Mirrors the table in ``docs/typed-io-contract.md`` cell by cell. The four
range shapes (cell / row / column / grid) crossed with the nine declared
types are the whole contract, so each combination is pinned here — both the
value it builds and the message it refuses with.
"""

import numpy as np
import pandas as pd
import pytest

from pyexcel.kernel import declared_types as dt


# Grids are row-major, matching arrow_io.decode_grid.
CELL = [[42]]
ROW = [["a", "b", "c"]]
COLUMN = [["hdr"], [1], [2]]
GRID = [["a", "b"], [1, 2], [3, 4]]


# -----------------------------------------------------------------------------
# Auto resolution
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "rows,cols,expected",
    [
        (1, 1, dt.SCALAR),
        (1, 5, dt.LIST),
        (5, 1, dt.LIST),
        (5, 3, dt.DATAFRAME),
        (0, 0, dt.SCALAR),
    ],
)
def test_resolve_auto(rows, cols, expected):
    assert dt.resolve_auto(rows, cols) == expected


def test_auto_is_resolved_by_build():
    assert dt.build(dt.AUTO, CELL, name="x") == 42
    assert dt.build(dt.AUTO, COLUMN, name="x") == ["hdr", 1, 2]
    assert isinstance(dt.build(dt.AUTO, GRID, name="x"), pd.DataFrame)


# -----------------------------------------------------------------------------
# DataFrame — first row is always the header
# -----------------------------------------------------------------------------


def test_dataframe_from_grid_takes_header_row():
    df = dt.build(dt.DATAFRAME, GRID, name="x")
    assert list(df.columns) == ["a", "b"]
    assert df.to_dict("list") == {"a": [1, 3], "b": [2, 4]}


def test_dataframe_from_column_names_the_single_column():
    df = dt.build(dt.DATAFRAME, COLUMN, name="x")
    assert list(df.columns) == ["hdr"]
    assert df["hdr"].tolist() == [1, 2]


def test_dataframe_from_cell_is_a_named_empty_column():
    """The documented one-cell case: the cell becomes the column name."""
    df = dt.build(dt.DATAFRAME, [["Total"]], name="x")
    assert list(df.columns) == ["Total"]
    assert len(df) == 0


def test_dataframe_from_row_has_columns_but_no_rows():
    df = dt.build(dt.DATAFRAME, ROW, name="x")
    assert list(df.columns) == ["a", "b", "c"]
    assert len(df) == 0


def test_blank_header_cell_falls_back_to_its_index():
    df = dt.build(dt.DATAFRAME, [["a", None, "  "], [1, 2, 3]], name="x")
    assert list(df.columns) == ["a", "1", "2"]


# -----------------------------------------------------------------------------
# Series
# -----------------------------------------------------------------------------


def test_series_from_column_is_named_by_the_top_cell():
    s = dt.build(dt.SERIES, COLUMN, name="x")
    assert s.name == "hdr"
    assert s.tolist() == [1, 2]


def test_series_from_row_is_named_by_the_first_cell():
    s = dt.build(dt.SERIES, ROW, name="x")
    assert s.name == "a"
    assert s.tolist() == ["b", "c"]


def test_series_from_grid_is_rejected():
    with pytest.raises(dt.TypeContractError) as exc:
        dt.build(dt.SERIES, GRID, name="Sales", address="Sheet1!A1:B3")
    message = str(exc.value)
    assert "Sales" in message
    assert "Sheet1!A1:B3" in message
    assert "3x2" in message
    assert "Series" in message


# -----------------------------------------------------------------------------
# Raw sequence types — every cell is data, nothing is consumed as a name
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "declared,factory", [(dt.LIST, list), (dt.TUPLE, tuple)]
)
def test_sequences_are_flat_for_a_row_or_column(declared, factory):
    assert dt.build(declared, ROW, name="x") == factory(["a", "b", "c"])
    assert dt.build(declared, COLUMN, name="x") == factory(["hdr", 1, 2])
    assert dt.build(declared, CELL, name="x") == factory([42])


@pytest.mark.parametrize(
    "declared,factory", [(dt.LIST, list), (dt.TUPLE, tuple)]
)
def test_sequences_nest_for_a_grid(declared, factory):
    built = dt.build(declared, GRID, name="x")
    assert built == factory([factory(["a", "b"]), factory([1, 2]), factory([3, 4])])


def test_set_flattens_and_dedupes():
    assert dt.build(dt.SET, [[1, 2], [2, 3]], name="x") == {1, 2, 3}
    assert dt.build(dt.SET, CELL, name="x") == {42}


@pytest.mark.parametrize(
    "grid,shape",
    [(CELL, (1,)), (ROW, (3,)), (COLUMN, (3,)), (GRID, (3, 2))],
)
def test_ndarray_matches_the_range_dimensions(grid, shape):
    assert dt.build(dt.NDARRAY, grid, name="x").shape == shape


def test_ndarray_is_a_real_numpy_array():
    assert isinstance(dt.build(dt.NDARRAY, GRID, name="x"), np.ndarray)


# -----------------------------------------------------------------------------
# Dict — two columns are pairs, wider is column-oriented
# -----------------------------------------------------------------------------


def test_dict_two_columns_reads_as_key_value_pairs():
    assert dt.build(dt.DICT, [["a", 1], ["b", 2]], name="x") == {"a": 1, "b": 2}


def test_dict_wider_than_two_columns_is_column_oriented():
    built = dt.build(dt.DICT, [["a", "b", "c"], [1, 2, 3], [4, 5, 6]], name="x")
    assert built == {"a": [1, 4], "b": [2, 5], "c": [3, 6]}


@pytest.mark.parametrize("grid", [CELL, COLUMN])
def test_dict_needs_at_least_two_columns(grid):
    with pytest.raises(dt.TypeContractError) as exc:
        dt.build(dt.DICT, grid, name="Rates", address="Sheet1!A1:A3")
    message = str(exc.value)
    assert "Rates" in message
    assert "at least 2 columns" in message
    assert "key -> value" in message


# -----------------------------------------------------------------------------
# Scalar — one cell only
# -----------------------------------------------------------------------------


def test_scalar_from_a_single_cell():
    assert dt.build(dt.SCALAR, CELL, name="x") == 42


@pytest.mark.parametrize("grid", [ROW, COLUMN, GRID])
def test_scalar_rejects_anything_larger_than_one_cell(grid):
    with pytest.raises(dt.TypeContractError) as exc:
        dt.build(dt.SCALAR, grid, name="Rate", address="Sheet1!A1:A3")
    message = str(exc.value)
    assert "Rate" in message
    assert "single cell" in message


def test_unknown_declared_type_names_the_valid_set():
    with pytest.raises(dt.TypeContractError) as exc:
        dt.build("frobnicate", CELL, name="x")
    assert "DataFrame" in str(exc.value)


# -----------------------------------------------------------------------------
# Output validation
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "declared,value",
    [
        (dt.DATAFRAME, pd.DataFrame({"a": [1]})),
        (dt.SERIES, pd.Series([1])),
        (dt.LIST, [1, 2]),
        (dt.TUPLE, (1, 2)),
        (dt.SET, {1, 2}),
        (dt.SET, frozenset({1})),
        (dt.DICT, {"a": 1}),
        (dt.NDARRAY, np.array([1])),
        (dt.SCALAR, 42),
        (dt.SCALAR, "text"),
        (dt.SCALAR, None),
    ],
)
def test_check_output_accepts_a_matching_value(declared, value):
    dt.check_output(declared, value, name="out")


@pytest.mark.parametrize(
    "declared,value",
    [
        (dt.DATAFRAME, [1, 2]),
        (dt.SERIES, pd.DataFrame({"a": [1]})),
        (dt.LIST, (1, 2)),
        (dt.TUPLE, [1, 2]),
        (dt.DICT, [1, 2]),
        (dt.SCALAR, [1]),
        (dt.SCALAR, pd.DataFrame({"a": [1]})),
        (dt.NDARRAY, [1, 2]),
    ],
)
def test_check_output_rejects_a_mismatch(declared, value):
    with pytest.raises(dt.TypeContractError) as exc:
        dt.check_output(declared, value, name="ProcessedSales")
    message = str(exc.value)
    assert "ProcessedSales" in message
    assert dt.display_name(declared) in message


@pytest.mark.parametrize("declared", [None, "", dt.AUTO])
def test_auto_output_is_never_enforced(declared):
    dt.check_output(declared, [1, 2, 3], name="out")
    dt.check_output(declared, pd.DataFrame({"a": [1]}), name="out")


def test_figures_are_exempt_from_output_enforcement():
    """A figure is converted at the encode door and ignores cell geometry,
    so enforcing a declared type against it would be a false positive."""

    class FakeFigure:
        pass

    FakeFigure.__module__ = "plotly.graph_objects"
    dt.check_output(dt.DATAFRAME, FakeFigure(), name="chart")
