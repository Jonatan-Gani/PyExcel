"""Tests for ``pyexcel.kernel.arrow_io``.

Covers shape-preserving roundtrip for the three first-class kernel
shapes (table / vector / scalar), Series and numpy bridging, edge cases
(empty, None, mixed types), decode of an externally produced Arrow
stream that has no PyExcel metadata, orientation peek, and the
TypeError contract for unsupported inputs.
"""

from __future__ import annotations

import datetime as dt
from decimal import Decimal

import numpy as np
import pandas as pd
import pyarrow as pa
import pyarrow.ipc as ipc
import pytest

from pyexcel.kernel.arrow_io import (
    Orientation,
    Shape,
    decode,
    decode_orientation,
    encode,
)


# -----------------------------------------------------------------------------
# DataFrame
# -----------------------------------------------------------------------------


def test_dataframe_roundtrip_preserves_columns_and_values():
    df = pd.DataFrame(
        {
            "i": [1, 2, 3],
            "f": [1.5, 2.5, 3.5],
            "s": ["a", "b", "c"],
            "b": [True, False, True],
        }
    )
    out = decode(encode(df))
    assert isinstance(out, pd.DataFrame)
    pd.testing.assert_frame_equal(out, df, check_dtype=False)


def test_dataframe_roundtrip_with_nulls():
    df = pd.DataFrame({"x": [1, None, 3], "y": ["a", None, "c"]})
    out = decode(encode(df))
    assert list(out.columns) == ["x", "y"]
    # The 'x' column comes back as float because of the NaN; that's pandas,
    # not us. What we care about is that the missing value survives.
    assert out["x"].isna().tolist() == [False, True, False]
    assert out["y"].isna().tolist() == [False, True, False]


def test_dataframe_empty():
    df = pd.DataFrame({"a": pd.Series([], dtype="int64"), "b": pd.Series([], dtype="object")})
    out = decode(encode(df))
    assert isinstance(out, pd.DataFrame)
    assert out.shape == (0, 2)
    assert list(out.columns) == ["a", "b"]


def test_dataframe_drops_default_rangeindex():
    df = pd.DataFrame({"x": [10, 20, 30]})
    # The default RangeIndex should not become a column on the wire.
    out = decode(encode(df))
    assert list(out.columns) == ["x"]
    assert "__index_level_0__" not in out.columns


# -----------------------------------------------------------------------------
# Vectors (list / tuple / Series / numpy 1-D)
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "value",
    [
        [1, 2, 3],
        [1.0, 2.5, 3.25],
        ["a", "b", "c"],
        [True, False, True],
        [1, None, 3],
        [],
    ],
)
def test_list_roundtrip(value):
    assert decode(encode(value)) == value


def test_tuple_decodes_to_list():
    # Tuples are accepted on encode for ergonomics, but the wire format
    # is "vector" — we don't carry container-flavour metadata, so the
    # canonical Python form coming back is list.
    assert decode(encode((1, 2, 3))) == [1, 2, 3]


def test_series_roundtrips_as_list():
    s = pd.Series([10, 20, 30], name="qty")
    out = decode(encode(s))
    assert out == [10, 20, 30]


def test_numpy_1d_roundtrips_as_list():
    arr = np.array([1, 2, 3], dtype=np.int32)
    out = decode(encode(arr))
    assert out == [1, 2, 3]


def test_numpy_2d_roundtrips_as_dataframe():
    arr = np.array([[1, 2], [3, 4], [5, 6]], dtype=np.int64)
    out = decode(encode(arr))
    assert isinstance(out, pd.DataFrame)
    assert out.shape == (3, 2)
    assert list(out.columns) == ["0", "1"]
    assert out.values.tolist() == [[1, 2], [3, 4], [5, 6]]


def test_numpy_3d_rejected():
    arr = np.zeros((2, 2, 2))
    with pytest.raises(TypeError, match="1-D or 2-D"):
        encode(arr)


# -----------------------------------------------------------------------------
# Scalars
# -----------------------------------------------------------------------------


@pytest.mark.parametrize(
    "value",
    [
        0,
        42,
        -7,
        3.14,
        "hello",
        "",
        True,
        False,
        None,
        b"binary",
        dt.date(2024, 1, 1),
        dt.datetime(2024, 1, 1, 12, 30, 45),
    ],
)
def test_scalar_roundtrip(value):
    assert decode(encode(value)) == value


def test_decimal_scalar_roundtrips_via_arrow():
    # Decimal goes through Arrow's decimal128 type — verify it survives.
    value = Decimal("123.45")
    out = decode(encode(value))
    assert out == value


def test_none_scalar_decodes_to_none():
    assert decode(encode(None)) is None


# -----------------------------------------------------------------------------
# Orientation hint
# -----------------------------------------------------------------------------


def test_default_orientation_is_column():
    assert decode_orientation(encode([1, 2, 3])) is Orientation.COLUMN


def test_explicit_row_orientation():
    buf = encode([1, 2, 3], orientation=Orientation.ROW)
    assert decode_orientation(buf) is Orientation.ROW


def test_orientation_is_none_for_tables():
    df = pd.DataFrame({"x": [1, 2]})
    assert decode_orientation(encode(df)) is None


def test_orientation_is_none_for_scalars():
    assert decode_orientation(encode(42)) is None


# -----------------------------------------------------------------------------
# Compatibility: external Arrow stream with no PyExcel metadata
# -----------------------------------------------------------------------------


def test_external_arrow_stream_decodes_as_dataframe():
    # Build an Arrow IPC stream the "vanilla" way, with no shape metadata.
    table = pa.table({"x": [1, 2, 3], "y": ["a", "b", "c"]})
    sink = pa.BufferOutputStream()
    with ipc.new_stream(sink, table.schema) as writer:
        writer.write_table(table)
    raw = bytes(sink.getvalue())

    out = decode(raw)
    assert isinstance(out, pd.DataFrame)
    assert list(out.columns) == ["x", "y"]
    assert out["x"].tolist() == [1, 2, 3]


# -----------------------------------------------------------------------------
# Errors
# -----------------------------------------------------------------------------


def test_unsupported_object_raises_typeerror():
    class Custom:
        pass

    with pytest.raises(TypeError):
        encode(Custom())


def test_unsupported_list_element_raises_typeerror():
    with pytest.raises(TypeError):
        encode([object(), object()])


def test_decode_garbage_bytes_raises():
    with pytest.raises(Exception):
        decode(b"this is not an arrow stream")


# -----------------------------------------------------------------------------
# Sanity: encoded output is non-empty bytes
# -----------------------------------------------------------------------------


def test_encode_returns_nonempty_bytes():
    out = encode(42)
    assert isinstance(out, bytes)
    assert len(out) > 0


def test_shape_enum_values_are_bytes():
    # The schema metadata is dict[bytes, bytes]; if these were str the
    # encode path would silently produce buffers that decode as DataFrame.
    assert isinstance(Shape.TABLE.value, bytes)
    assert isinstance(Shape.VECTOR.value, bytes)
    assert isinstance(Shape.SCALAR.value, bytes)
