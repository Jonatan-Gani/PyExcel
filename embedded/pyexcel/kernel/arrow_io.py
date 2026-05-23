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
from typing import Any, Tuple

import pyarrow as pa
import pyarrow.ipc as ipc

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


class Shape(bytes, enum.Enum):
    """The high-level shape of a value flowing across the kernel boundary.

    Stored as raw bytes because Arrow schema metadata is a ``dict[bytes, bytes]``.
    """

    TABLE = b"table"
    VECTOR = b"vector"
    SCALAR = b"scalar"


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

    if isinstance(value, (list, tuple)):
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

    if shape_bytes == Shape.SCALAR.value:
        if table.num_columns == 0 or table.num_rows == 0:
            return None
        return table.column(0)[0].as_py()

    if shape_bytes == Shape.VECTOR.value:
        if table.num_columns == 0:
            return []
        return table.column(0).to_pylist()

    # Default: shape=table (or any unrecognised marker we treat as table).
    if _HAS_PANDAS:
        return table.to_pandas()
    return table


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
