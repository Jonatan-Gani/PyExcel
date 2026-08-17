"""
Arrow IPC marshalling for the PyExcel v2 kernel data plane.

The framing layer carries opaque payloads; this module decides what goes in
them. Every value crossing the kernel boundary — Excel range as input,
Python return value as output — is serialised as a single Arrow IPC
*stream* (one schema, one or more record batches, no footer).

To preserve the caller's original Python shape across a format that only
knows tables, we piggyback on Arrow schema metadata:

    pyexcel-shape       = "table" | "vector" | "scalar"
    pyexcel-orientation = "row" | "column"   (vectors only)

Decoding inverts the encoding so a roundtrip is shape-preserving:

    table   → pandas.DataFrame
    vector  → list (orientation is carried so the host can spill correctly)
    scalar  → the unwrapped Python value (int / float / str / bool / None / …)

Design rules:

* **One way in, one way out.** A single ``encode(value) -> bytes`` and
  ``decode(buf) -> Any`` keep the contract small. Anything Arrow can
  represent works; anything else raises ``TypeError`` at encode time.
* **Stream, not file.** ``pyarrow.ipc.new_stream`` produces a forward-only
  stream; the C# counterpart reads it with ``ArrowStreamReader``. The file
  format (random-access footer) is intentionally not used.
* **Default to table.** A payload with no ``pyexcel-shape`` metadata
  decodes as a DataFrame. That keeps us forward-compatible with payloads
  produced by external Arrow writers.
* **No silent coercions.** Encoding fails loudly for unsupported types
  rather than dropping data or guessing.
"""

from __future__ import annotations

import enum
from typing import Any, List, Tuple

import pyarrow as pa
import pyarrow.ipc as ipc

from .chart import convert_figure
from .types import ChartImage, ChartSpec, Formula

try:
    import pandas as pd

    _HAS_PANDAS = True
except ImportError:  # pragma: no cover — pandas is a hard kernel dep
    _HAS_PANDAS = False

try:
    import numpy as np

    _HAS_NUMPY = True
except ImportError:  # pragma: no cover — numpy is a hard kernel dep
    _HAS_NUMPY = False


_META_SHAPE = b"pyexcel-shape"
_META_ORIENT = b"pyexcel-orientation"

# Field-level metadata key. Marks an Arrow string column whose cell values
# are formula source rather than literals — the host writes those via
# Range.Formula instead of Range.Value2 so Excel evaluates them.
_FIELD_META_CELL_TYPE = b"pyexcel-cell-type"
_CELL_TYPE_FORMULA = b"formula"

# Field-level metadata key for image payloads: the rendered format of the
# binary column ("svg" or "png"). Only present on shape=image buffers.
_FIELD_META_IMAGE_FORMAT = b"pyexcel-image-format"


class Shape(bytes, enum.Enum):
    """The high-level shape of a value flowing across the kernel boundary.

    Stored as raw bytes because Arrow schema metadata is a ``dict[bytes, bytes]``.
    """

    TABLE = b"table"
    VECTOR = b"vector"
    SCALAR = b"scalar"
    # Chart spec JSON (string scalar) — the host builds a native Excel chart.
    CHART = b"chart"
    # Rendered figure image (binary scalar) — the host embeds a picture.
    IMAGE = b"image"


class Orientation(bytes, enum.Enum):
    """Vector orientation hint for the host's spill direction.

    Has no effect on the Python side (vectors decode to ``list`` either
    way); it's purely advisory metadata for the C# host.
    """

    ROW = b"row"
    COLUMN = b"column"


# -----------------------------------------------------------------------------
# Encode
# -----------------------------------------------------------------------------


def encode(value: Any, *, orientation: Orientation = Orientation.COLUMN) -> bytes:
    """Serialise a Python value to an Arrow IPC stream.

    Args:
        value: A ``pandas.DataFrame``, ``pandas.Series``, ``list``/``tuple``,
            1-D ``numpy.ndarray``, or any scalar Arrow can represent.
        orientation: Spill direction hint for vector inputs. Ignored for
            tables and scalars.

    Returns:
        The encoded Arrow IPC stream as ``bytes``.

    Raises:
        TypeError: ``value`` is not one of the supported shapes.
    """
    # Plotly figure → ChartSpec, Matplotlib figure → ChartImage; everything
    # else passes through untouched. Conversion happens at the encode door
    # so user transform() functions can return figures directly.
    value = convert_figure(value)

    table, shape = _value_to_table(value)

    metadata = dict(table.schema.metadata) if table.schema.metadata else {}
    metadata[_META_SHAPE] = shape.value
    if shape is Shape.VECTOR:
        metadata[_META_ORIENT] = orientation.value
    table = table.replace_schema_metadata(metadata)

    sink = pa.BufferOutputStream()
    with ipc.new_stream(sink, table.schema) as writer:
        writer.write_table(table)
    return bytes(sink.getvalue())


def _value_to_table(value: Any) -> Tuple[pa.Table, Shape]:
    """Coerce ``value`` into ``(Arrow Table, original-shape)``."""
    if _HAS_PANDAS and isinstance(value, pd.DataFrame):
        return pa.Table.from_pandas(value, preserve_index=False), Shape.TABLE

    if _HAS_PANDAS and isinstance(value, pd.Series):
        name = value.name if value.name is not None else "0"
        return (
            pa.Table.from_arrays([pa.Array.from_pandas(value)], names=[str(name)]),
            Shape.VECTOR,
        )

    if _HAS_NUMPY and isinstance(value, np.ndarray):
        if value.ndim == 1:
            return pa.Table.from_arrays([pa.array(value)], names=["0"]), Shape.VECTOR
        if value.ndim == 2:
            # 2-D arrays become unnamed tables; column labels are positional
            # strings so the receiver gets a deterministic schema.
            cols = [pa.array(value[:, i]) for i in range(value.shape[1])]
            names = [str(i) for i in range(value.shape[1])]
            return pa.Table.from_arrays(cols, names=names), Shape.TABLE
        raise TypeError(
            f"numpy arrays must be 1-D or 2-D for kernel transport, got {value.ndim}-D"
        )

    if isinstance(value, ChartSpec):
        # Chart spec: 1×1 string table; the CHART shape marker is what the
        # host dispatches on, no field metadata needed.
        arr = pa.array([value.json], type=pa.string())
        return pa.Table.from_arrays([arr], names=["0"]), Shape.CHART

    if isinstance(value, ChartImage):
        # Image: 1×1 binary table with the rendered format on the field so
        # the host knows the file extension to embed with.
        field = pa.field("0", pa.binary()).with_metadata(
            {_FIELD_META_IMAGE_FORMAT: value.format.encode("ascii")}
        )
        arr = pa.array([value.data], type=pa.binary())
        return pa.Table.from_arrays([arr], schema=pa.schema([field])), Shape.IMAGE

    if isinstance(value, Formula):
        # Scalar formula: 1×1 string table with the formula field marker so
        # decode wraps the cell back into a Formula, and the host writes it
        # via Range.Formula instead of Range.Value2.
        field = pa.field("0", pa.string()).with_metadata(
            {_FIELD_META_CELL_TYPE: _CELL_TYPE_FORMULA}
        )
        arr = pa.array([value.text], type=pa.string())
        return pa.Table.from_arrays([arr], schema=pa.schema([field])), Shape.SCALAR

    if isinstance(value, (list, tuple)):
        # All-Formula list → formula-marked string vector. A mixed list
        # (some Formula, some not) is rejected as ambiguous: today's wire
        # marker is per-column, so we can't carry per-cell type info in
        # a 1-D vector without a much wider redesign.
        if value and all(isinstance(v, Formula) for v in value):
            field = pa.field("0", pa.string()).with_metadata(
                {_FIELD_META_CELL_TYPE: _CELL_TYPE_FORMULA}
            )
            arr = pa.array([v.text for v in value], type=pa.string())
            return pa.Table.from_arrays([arr], schema=pa.schema([field])), Shape.VECTOR
        if any(isinstance(v, Formula) for v in value):
            raise TypeError(
                "list/tuple may not mix Formula with non-Formula entries; "
                "per-cell formula marking isn't supported in 1-D payloads"
            )
        try:
            arr = pa.array(value)
        except (pa.ArrowInvalid, pa.ArrowTypeError) as exc:
            raise TypeError(
                f"could not encode {type(value).__name__} as a 1-D Arrow array: {exc}"
            ) from exc
        return pa.Table.from_arrays([arr], names=["0"]), Shape.VECTOR

    # Anything else: treat as a scalar wrapped in a 1×1 table. Includes
    # int, float, str, bool, bytes, None, datetime, Decimal, … — whatever
    # Arrow can infer a type for from a length-1 list.
    try:
        arr = pa.array([value])
    except (pa.ArrowInvalid, pa.ArrowTypeError) as exc:
        raise TypeError(
            f"unsupported scalar type {type(value).__name__}: {exc}"
        ) from exc
    return pa.Table.from_arrays([arr], names=["0"]), Shape.SCALAR


# -----------------------------------------------------------------------------
# Decode
# -----------------------------------------------------------------------------


def decode(buf: bytes) -> Any:
    """Deserialise an Arrow IPC stream back to its original Python shape.

    Shape is recovered from schema metadata written by :func:`encode`.
    Buffers produced by external Arrow writers (no ``pyexcel-shape`` key)
    decode as a DataFrame, matching the most common case.

    Raises:
        pyarrow.ArrowInvalid: ``buf`` is not a valid Arrow IPC stream.
    """
    reader = ipc.open_stream(pa.BufferReader(buf))
    table = reader.read_all()
    metadata = table.schema.metadata or {}
    shape_bytes = metadata.get(_META_SHAPE, Shape.TABLE.value)

    if shape_bytes == Shape.CHART.value:
        if table.num_columns == 0 or table.num_rows == 0:
            raise ValueError("chart-shaped buffer carries no spec cell")
        return ChartSpec(table.column(0)[0].as_py())

    if shape_bytes == Shape.IMAGE.value:
        if table.num_columns == 0 or table.num_rows == 0:
            raise ValueError("image-shaped buffer carries no data cell")
        field_md = table.schema.field(0).metadata or {}
        fmt = field_md.get(_FIELD_META_IMAGE_FORMAT, b"png").decode("ascii")
        return ChartImage(table.column(0)[0].as_py(), fmt)

    if shape_bytes == Shape.SCALAR.value:
        if table.num_columns == 0 or table.num_rows == 0:
            return None
        cell = table.column(0)[0].as_py()
        if _is_formula_field(table.schema.field(0)):
            return None if cell is None else Formula(cell)
        return cell

    if shape_bytes == Shape.VECTOR.value:
        if table.num_columns == 0:
            return []
        values = table.column(0).to_pylist()
        if _is_formula_field(table.schema.field(0)):
            return [None if v is None else Formula(v) for v in values]
        return values

    # Default: shape=table (or any unrecognised marker we treat as table).
    return _table_to_python(table)


def decode_grid(buf: bytes) -> List[List[Any]]:
    """Deserialise a buffer to a raw row-major grid of cell values.

    This is the input path for the typed contract. Where :func:`decode`
    recovers whatever shape the *producer* intended, ``decode_grid``
    deliberately recovers nothing: it hands back the cells exactly as they
    sit in the sheet so :mod:`pyexcel.kernel.declared_types` can build the
    type the *user declared*. Shape metadata is ignored on purpose.

    A scalar buffer yields ``[[value]]`` and a vector yields a single row or
    column, so every range shape arrives as a rectangle and the coercion
    matrix has one uniform input format.

    Raises:
        pyarrow.ArrowInvalid: ``buf`` is not a valid Arrow IPC stream.
    """
    reader = ipc.open_stream(pa.BufferReader(buf))
    table = reader.read_all()
    if table.num_columns == 0:
        return []

    columns = [table.column(c).to_pylist() for c in range(table.num_columns)]

    # A vector encoded as a row still arrives as one Arrow column, because
    # Arrow has no notion of orientation; the orientation metadata is what
    # tells us to lay it back out across rather than down.
    metadata = table.schema.metadata or {}
    if metadata.get(_META_SHAPE) == Shape.VECTOR.value:
        if metadata.get(_META_ORIENT) == Orientation.ROW.value:
            return [list(columns[0])]

    rows = len(columns[0])
    return [[column[r] for column in columns] for r in range(rows)]


def _is_formula_field(field: pa.Field) -> bool:
    """Whether an Arrow field carries the formula cell-type marker."""
    md = field.metadata or {}
    return md.get(_FIELD_META_CELL_TYPE) == _CELL_TYPE_FORMULA


def _table_to_python(table: pa.Table) -> Any:
    """Convert a table-shaped buffer back to a Python object.

    If any field carries the formula marker, the table can't go through
    ``table.to_pandas()`` directly — pandas would silently strip the
    field metadata and turn formula columns into plain strings. Build
    the DataFrame column-by-column, wrapping marked columns into a
    pandas object Series of :class:`Formula` instances. If pandas isn't
    available, return the raw Arrow ``Table``.
    """
    if not _HAS_PANDAS:
        return table

    formula_indices = [
        i for i in range(table.num_columns)
        if _is_formula_field(table.schema.field(i))
    ]
    if not formula_indices:
        return table.to_pandas()

    # Mixed-column table with at least one formula column. Build each
    # column as a Series with the right dtype, then assemble into a
    # DataFrame so the result is interchangeable with `table.to_pandas()`
    # for user code that does df[col].
    formula_index_set = set(formula_indices)
    columns: List[pd.Series] = []
    for i in range(table.num_columns):
        name = table.schema.field(i).name
        if i in formula_index_set:
            vals = table.column(i).to_pylist()
            columns.append(
                pd.Series(
                    [None if v is None else Formula(v) for v in vals],
                    name=name,
                    dtype="object",
                )
            )
        else:
            # One-column slice → to_pandas() yields a DataFrame; pull the
            # Series out by position so we don't depend on the name being
            # unique across the original table.
            series = table.select([i]).to_pandas().iloc[:, 0]
            series.name = name
            columns.append(series)
    return pd.concat(columns, axis=1)


def decode_orientation(buf: bytes) -> Orientation | None:
    """Peek at a vector buffer's orientation hint without materialising the data.

    Returns ``None`` for non-vector payloads or buffers missing the hint.
    The host uses this to decide spill direction before allocating cells.
    """
    reader = ipc.open_stream(pa.BufferReader(buf))
    schema = reader.schema
    metadata = schema.metadata or {}
    if metadata.get(_META_SHAPE) != Shape.VECTOR.value:
        return None
    raw = metadata.get(_META_ORIENT)
    if raw is None:
        return None
    try:
        return Orientation(raw)
    except ValueError:
        return None
