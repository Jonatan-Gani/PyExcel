"""The declared-type vocabulary and coercion matrix.

This module is the kernel half of the typed I/O contract specified in
``docs/typed-io-contract.md``. The host resolves each binding to a concrete
type (it is the side that has measured the range) and sends the cells as a
raw R x C grid; everything here turns that grid into the declared Python
type, or explains precisely why it cannot.

The wire names below are mirrored by ``PyExcelTypes.WireName`` in
``src/PyExcel.State/PyExcelType.cs``. Changing one requires changing the
other.

Two directions, deliberately asymmetric:

* **Inputs are constructed.** Excel hands us nothing but cells, so a
  declared input type is an instruction for what to build.
* **Outputs are validated.** Python hands back a real object, so a declared
  output type is an assertion about what came back. Nothing is silently
  converted on the way out.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Sequence

try:  # pragma: no cover - exercised by the absence path only
    import pandas as pd

    _HAS_PANDAS = True
except ImportError:  # pragma: no cover
    pd = None  # type: ignore[assignment]
    _HAS_PANDAS = False

try:  # pragma: no cover
    import numpy as np

    _HAS_NUMPY = True
except ImportError:  # pragma: no cover
    np = None  # type: ignore[assignment]
    _HAS_NUMPY = False


# -----------------------------------------------------------------------------
# Vocabulary
# -----------------------------------------------------------------------------

AUTO = "auto"
DATAFRAME = "dataframe"
SERIES = "series"
LIST = "list"
TUPLE = "tuple"
SET = "set"
DICT = "dict"
NDARRAY = "ndarray"
SCALAR = "scalar"

#: Every declared type the wire may carry, in the dropdown's order.
ALL_TYPES = (AUTO, DATAFRAME, SERIES, LIST, TUPLE, SET, DICT, NDARRAY, SCALAR)

#: Display names used in error messages, matching the form's type box.
_DISPLAY = {
    AUTO: "Auto",
    DATAFRAME: "DataFrame",
    SERIES: "Series",
    LIST: "List",
    TUPLE: "Tuple",
    SET: "Set",
    DICT: "Dict",
    NDARRAY: "NDArray",
    SCALAR: "Scalar",
}


def display_name(declared: str) -> str:
    """Human-facing label for a wire type name."""
    return _DISPLAY.get(declared, declared)


class TypeContractError(Exception):
    """A declared type could not be satisfied.

    The worker converts this into a ``BadInput`` (construction) or
    ``BadReturnType`` (validation) job error. The message is written for the
    user who configured the binding, not for a developer reading a stack
    trace: it names the binding, the range, the measured dimensions, the
    declared type, and the way out.
    """


def resolve_auto(rows: int, columns: int) -> str:
    """Resolve ``auto`` from measured dimensions.

    Mirrors ``PyExcelTypes.ResolveAuto``. The host normally resolves this
    before the request is sent; the kernel keeps its own copy so a direct
    kernel-client caller can pass ``auto`` and still get the documented
    default.
    """
    if rows <= 1 and columns <= 1:
        return SCALAR
    if rows <= 1 or columns <= 1:
        return LIST
    return DATAFRAME


# -----------------------------------------------------------------------------
# Input construction
# -----------------------------------------------------------------------------


def _describe(name: str, address: str, rows: int, columns: int) -> str:
    """The '<binding> (<range>, RxC)' prefix every error message opens with."""
    where = f" ({address}, {rows}x{columns})" if address else f" ({rows}x{columns})"
    return f"input '{name}'{where}"


def _require(condition: bool, message: str) -> None:
    if not condition:
        raise TypeContractError(message)


def _grid_shape(grid: Sequence[Sequence[Any]]) -> tuple:
    rows = len(grid)
    columns = len(grid[0]) if rows else 0
    return rows, columns


def build(
    declared: str,
    grid: Sequence[Sequence[Any]],
    *,
    name: str,
    address: str = "",
) -> Any:
    """Construct ``declared`` from a raw row-major ``grid`` of cell values.

    Args:
        declared: A wire type name. ``auto`` is resolved from the grid's
            dimensions, so a caller that has not resolved it still works.
        grid: Row-major cell values, as produced by ``arrow_io.decode_grid``.
        name: The binding's name, used in error messages.
        address: The binding's range text, used in error messages.

    Raises:
        TypeContractError: The grid cannot satisfy the declared type.
    """
    rows, columns = _grid_shape(grid)
    if declared == AUTO:
        declared = resolve_auto(rows, columns)

    where = _describe(name, address, rows, columns)

    if declared == DATAFRAME:
        return _build_dataframe(grid, where)
    if declared == SERIES:
        return _build_series(grid, rows, columns, where)
    if declared == LIST:
        return _build_sequence(grid, rows, columns, list)
    if declared == TUPLE:
        return _build_tuple(grid, rows, columns)
    if declared == SET:
        return {cell for row in grid for cell in row}
    if declared == NDARRAY:
        return _build_ndarray(grid, rows, columns, where)
    if declared == DICT:
        return _build_dict(grid, rows, columns, where)
    if declared == SCALAR:
        _require(
            rows == 1 and columns == 1,
            f"{where}: declared type Scalar needs a single cell, but the range is "
            f"{rows}x{columns}. Select one cell, or declare List, DataFrame or NDArray.",
        )
        return grid[0][0]

    raise TypeContractError(
        f"{where}: unknown declared type '{declared}'. Expected one of: "
        + ", ".join(display_name(t) for t in ALL_TYPES)
    )


def _build_dataframe(grid: Sequence[Sequence[Any]], where: str) -> Any:
    """First row supplies column names; every remaining row is data.

    One rule covers all four range shapes. A single cell yields a
    zero-row frame whose only column is named by that cell, which is the
    documented behaviour.
    """
    _require(
        _HAS_PANDAS,
        f"{where}: declared type DataFrame needs pandas, which is not installed "
        "in this project's environment. Re-run Setup from the ribbon's Enable button.",
    )
    if not grid:
        return pd.DataFrame()

    header = [_as_column_name(cell, i) for i, cell in enumerate(grid[0])]
    data = [list(row) for row in grid[1:]]
    return pd.DataFrame(data, columns=header)


def _as_column_name(cell: Any, index: int) -> str:
    """Column names are strings; an empty header cell falls back to its index."""
    if cell is None:
        return str(index)
    text = str(cell).strip()
    return text if text else str(index)


def _build_series(
    grid: Sequence[Sequence[Any]], rows: int, columns: int, where: str
) -> Any:
    _require(
        _HAS_PANDAS,
        f"{where}: declared type Series needs pandas, which is not installed "
        "in this project's environment. Re-run Setup from the ribbon's Enable button.",
    )
    _require(
        rows <= 1 or columns <= 1,
        f"{where}: declared type Series needs a single row or column, but the range "
        f"is {rows}x{columns}. Use DataFrame, or narrow the selection to one row or column.",
    )
    if not grid:
        return pd.Series(dtype="object")

    series_name = _as_column_name(grid[0][0], 0)
    if columns == 1:
        values = [row[0] for row in grid[1:]]
    else:
        values = list(grid[0][1:])
    return pd.Series(values, name=series_name)


def _build_sequence(
    grid: Sequence[Sequence[Any]], rows: int, columns: int, factory
) -> Any:
    """Flat for a single row or column, nested per row for a real grid."""
    if rows == 1:
        return factory(grid[0])
    if columns == 1:
        return factory(row[0] for row in grid)
    return factory(factory(row) for row in grid)


def _build_tuple(grid: Sequence[Sequence[Any]], rows: int, columns: int) -> Any:
    if rows == 1:
        return tuple(grid[0])
    if columns == 1:
        return tuple(row[0] for row in grid)
    return tuple(tuple(row) for row in grid)


def _build_ndarray(
    grid: Sequence[Sequence[Any]], rows: int, columns: int, where: str
) -> Any:
    _require(
        _HAS_NUMPY,
        f"{where}: declared type NDArray needs numpy, which is not installed "
        "in this project's environment. Re-run Setup from the ribbon's Enable button.",
    )
    if rows == 1:
        return np.array(list(grid[0]))
    if columns == 1:
        return np.array([row[0] for row in grid])
    return np.array([list(row) for row in grid])


def _build_dict(
    grid: Sequence[Sequence[Any]], rows: int, columns: int, where: str
) -> Dict[Any, Any]:
    """Two columns read as key -> value; three or more as column-oriented lists."""
    _require(
        columns >= 2,
        f"{where}: declared type Dict needs at least 2 columns, but the range is "
        f"{rows}x{columns}. Two columns are read as key -> value pairs; three or "
        "more as column-oriented lists keyed by the header row.",
    )
    if columns == 2:
        return {row[0]: row[1] for row in grid}

    header = [_as_column_name(cell, i) for i, cell in enumerate(grid[0])]
    return {
        key: [row[j] for row in grid[1:]]
        for j, key in enumerate(header)
    }


# -----------------------------------------------------------------------------
# Output validation
# -----------------------------------------------------------------------------


def _python_type_name(value: Any) -> str:
    """A user-facing name for what the script actually returned."""
    if _HAS_PANDAS:
        if isinstance(value, pd.DataFrame):
            return "DataFrame"
        if isinstance(value, pd.Series):
            return "Series"
    if _HAS_NUMPY and isinstance(value, np.ndarray):
        return "NDArray"
    return type(value).__name__


def _matches(declared: str, value: Any) -> bool:
    if declared == DATAFRAME:
        return _HAS_PANDAS and isinstance(value, pd.DataFrame)
    if declared == SERIES:
        return _HAS_PANDAS and isinstance(value, pd.Series)
    if declared == NDARRAY:
        return _HAS_NUMPY and isinstance(value, np.ndarray)
    if declared == LIST:
        return isinstance(value, list)
    if declared == TUPLE:
        return isinstance(value, tuple)
    if declared == SET:
        return isinstance(value, (set, frozenset))
    if declared == DICT:
        return isinstance(value, dict)
    if declared == SCALAR:
        return not isinstance(value, (list, tuple, set, frozenset, dict)) and not (
            (_HAS_PANDAS and isinstance(value, (pd.DataFrame, pd.Series)))
            or (_HAS_NUMPY and isinstance(value, np.ndarray))
        )
    return True


def check_output(declared: Optional[str], value: Any, *, name: str) -> None:
    """Assert a returned value satisfies its binding's declared type.

    ``auto`` and ``None`` mean "do not enforce" — the loose behaviour that
    keeps iterating on a script comfortable. Figures are always exempt: they
    are converted at the encode door and ignore cell geometry entirely, so
    there is nothing to validate them against.

    Raises:
        TypeContractError: The value does not match the declared type.
    """
    if not declared or declared == AUTO:
        return
    if _is_figure(value):
        return
    if _matches(declared, value):
        return

    raise TypeContractError(
        f"output '{name}': declared type {display_name(declared)}, but transform() "
        f"returned {_python_type_name(value)}."
    )


def _is_figure(value: Any) -> bool:
    """Whether a value is a chart/image figure, without importing plotting libs.

    Checked by module path rather than isinstance so the kernel does not drag
    plotly or matplotlib into a run that has no charts in it.
    """
    module = type(value).__module__ or ""
    return module.startswith("plotly") or module.startswith("matplotlib")
