using System;
using System.Collections.Generic;
using System.IO;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using Apache.Arrow.Types;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// C# roundtrip tests for <see cref="ArrowMarshal"/>. End-to-end agreement
/// with <c>arrow_io.py</c> is implicitly tested by the kernel integration
/// tests (KernelClientTests) once <c>=PY.RUN</c> lands — those tests are
/// the real cross-language conformance check, just like
/// CrossLanguageVectorsTests is for framing.
/// </summary>
public class ArrowMarshalTests
{
    // -------------------------------------------------------------------------
    // Shape peek
    // -------------------------------------------------------------------------

    [Fact]
    public void PeekShape_Table_ReportsTable()
    {
        var buf = ArrowMarshal.EncodeTable(new object?[,] { { 1.0, "a" } });
        var (shape, orientation) = ArrowMarshal.PeekShape(buf);
        Assert.Equal(ArrowShape.Table, shape);
        Assert.Null(orientation);
    }

    [Fact]
    public void PeekShape_Vector_ReportsVectorAndOrientation()
    {
        var buf = ArrowMarshal.EncodeVector(new object?[] { 1.0, 2.0 }, ArrowOrientation.Row);
        var (shape, orientation) = ArrowMarshal.PeekShape(buf);
        Assert.Equal(ArrowShape.Vector, shape);
        Assert.Equal(ArrowOrientation.Row, orientation);
    }

    [Fact]
    public void PeekShape_VectorDefaultOrientation_IsColumn()
    {
        var buf = ArrowMarshal.EncodeVector(new object?[] { 1.0, 2.0 });
        var (_, orientation) = ArrowMarshal.PeekShape(buf);
        Assert.Equal(ArrowOrientation.Column, orientation);
    }

    [Fact]
    public void PeekShape_Scalar_ReportsScalar()
    {
        var buf = ArrowMarshal.EncodeScalar(42.0);
        var (shape, orientation) = ArrowMarshal.PeekShape(buf);
        Assert.Equal(ArrowShape.Scalar, shape);
        Assert.Null(orientation);
    }

    // -------------------------------------------------------------------------
    // Table roundtrip
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeTable_AllDouble_RoundTrips()
    {
        var input = new object?[,]
        {
            { 1.0, 2.0, 3.0 },
            { 4.0, 5.0, 6.0 },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_MixedColumnTypes_RoundTripsWithPerColumnInference()
    {
        var input = new object?[,]
        {
            { 1.0, "a", true },
            { 2.0, "b", false },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_PreservesNulls()
    {
        var input = new object?[,]
        {
            { 1.0, null, "z" },
            { null, 5.0, null },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_AllNullColumn_StaysNullOnRoundtrip()
    {
        var input = new object?[,]
        {
            { 1.0, null },
            { 2.0, null },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        Assert.Null(decoded[0, 1]);
        Assert.Null(decoded[1, 1]);
        Assert.Equal(1.0, decoded[0, 0]);
        Assert.Equal(2.0, decoded[1, 0]);
    }

    [Fact]
    public void EncodeTable_MixedTypeColumn_FallsBackToString()
    {
        // Column 0 has a number and a string — no common primitive type,
        // so the encoder stringifies every value.
        var input = new object?[,]
        {
            { 1.0 },
            { "two" },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        Assert.Equal("1", decoded[0, 0]);
        Assert.Equal("two", decoded[1, 0]);
    }

    [Fact]
    public void EncodeTable_HonoursColumnNames()
    {
        var input = new object?[,] { { 1.0, "a" } };
        var buf = ArrowMarshal.EncodeTable(input, new[] { "qty", "label" });

        // Reading the header back requires the lower-level peek; the
        // decode here just confirms the values aren't reordered.
        var decoded = (object?[,])ArrowMarshal.Decode(buf)!;
        AssertTablesEqual(input, decoded);
    }

    // -------------------------------------------------------------------------
    // Vector roundtrip
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeVector_Doubles_RoundTrips()
    {
        var input = new object?[] { 1.0, 2.5, 3.0 };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(input, decoded);
    }

    [Fact]
    public void EncodeVector_Strings_RoundTrips()
    {
        var input = new object?[] { "a", "bb", "ccc" };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(input, decoded);
    }

    [Fact]
    public void EncodeVector_WithNulls_RoundTrips()
    {
        var input = new object?[] { 1.0, null, 3.0 };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(input, decoded);
    }

    [Fact]
    public void EncodeVector_AllNull_RoundTrips()
    {
        var input = new object?[] { null, null, null };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(3, decoded.Length);
        Assert.All(decoded, v => Assert.Null(v));
    }

    [Fact]
    public void EncodeVector_Bools_RoundTrips()
    {
        var input = new object?[] { true, false, true };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(input, decoded);
    }

    // -------------------------------------------------------------------------
    // Scalar roundtrip
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(0.0)]
    [InlineData(3.14)]
    [InlineData(-7.5)]
    public void EncodeScalar_Double_RoundTrips(double value)
    {
        var decoded = ArrowMarshal.Decode(ArrowMarshal.EncodeScalar(value));
        Assert.Equal(value, decoded);
    }

    [Theory]
    [InlineData("")]
    [InlineData("hello")]
    [InlineData("üñîçødé")]
    public void EncodeScalar_String_RoundTrips(string value)
    {
        var decoded = ArrowMarshal.Decode(ArrowMarshal.EncodeScalar(value));
        Assert.Equal(value, decoded);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EncodeScalar_Bool_RoundTrips(bool value)
    {
        var decoded = ArrowMarshal.Decode(ArrowMarshal.EncodeScalar(value));
        Assert.Equal(value, decoded);
    }

    [Fact]
    public void EncodeScalar_Null_RoundTripsToNull()
    {
        var decoded = ArrowMarshal.Decode(ArrowMarshal.EncodeScalar(null));
        Assert.Null(decoded);
    }

    [Fact]
    public void EncodeScalar_IntegerCoercesToDouble()
    {
        // Int → numeric-path → double. This matches Excel's "everything
        // numeric is a double" model so the kernel-side pandas sees a
        // consistent dtype.
        var decoded = ArrowMarshal.Decode(ArrowMarshal.EncodeScalar(42));
        Assert.Equal(42.0, decoded);
    }

    // -------------------------------------------------------------------------
    // Date / timestamp decoding
    //
    // Cells of these Arrow types arrive from pyarrow when a Python script
    // returns a datetime.date / datetime.datetime / numpy.datetime64 — the
    // C# encoder never produces them today (Excel hands us doubles for date
    // cells via Value2). So coverage is decode-only, and the buffers are
    // built straight from Apache.Arrow's API rather than via EncodeTable.
    // -------------------------------------------------------------------------

    [Fact]
    public void Decode_TimestampMicrosecond_YieldsNaiveDateTime()
    {
        // pyarrow's default for datetime.datetime is timestamp[us] with no
        // timezone. Build the equivalent buffer in-test and verify the
        // decoder lands on a naive DateTime.
        var expected = new DateTime(2024, 1, 15, 9, 30, 45, DateTimeKind.Unspecified);
        var microsSinceEpoch = (expected.Ticks - new DateTime(1970, 1, 1).Ticks) / 10L;
        var buf = BuildTimestampScalarBuffer(microsSinceEpoch, TimeUnit.Microsecond);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        Assert.Equal(expected, dt);
        Assert.Equal(DateTimeKind.Unspecified, dt.Kind);
    }

    [Fact]
    public void Decode_TimestampMillisecond_RoundsToNearestMs()
    {
        // External Arrow writers may use timestamp[ms] (Java/Spark default).
        // 1700000000000 ms after epoch is 2023-11-14 22:13:20 UTC; the
        // decoder treats it as naive (no timezone applied).
        var buf = BuildTimestampScalarBuffer(1_700_000_000_000L, TimeUnit.Millisecond);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        Assert.Equal(new DateTime(2023, 11, 14, 22, 13, 20, DateTimeKind.Unspecified), dt);
    }

    [Fact]
    public void Decode_TimestampSecond_RoundTrips()
    {
        // 1700000000 seconds after epoch.
        var buf = BuildTimestampScalarBuffer(1_700_000_000L, TimeUnit.Second);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        Assert.Equal(new DateTime(2023, 11, 14, 22, 13, 20, DateTimeKind.Unspecified), dt);
    }

    [Fact]
    public void Decode_TimestampNanosecond_TruncatesToTickPrecision()
    {
        // DateTime resolution is 100 ns, so 1234 ns past a tick boundary
        // truncates to a tick boundary. 1_700_000_000_000_001_234 ns
        // = 1700000000000001 ticks + 234 ns remainder; the value/100
        // division drops the remainder.
        var buf = BuildTimestampScalarBuffer(1_700_000_000_000_001_234L, TimeUnit.Nanosecond);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        // Expected: epoch + 17_000_000_000_000_012 ticks (= 1.7e15 ns / 100).
        var expected = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Unspecified)
            .AddTicks(17_000_000_000_000_012L);
        Assert.Equal(expected, dt);
    }

    [Fact]
    public void Decode_Date32Array_YieldsMidnightDateTime()
    {
        // pyarrow encodes datetime.date as date32 (days since 1970-01-01).
        var expected = new DateTime(2024, 1, 15, 0, 0, 0, DateTimeKind.Unspecified);
        var days = (int)((expected - new DateTime(1970, 1, 1)).TotalDays);
        var buf = BuildDate32ScalarBuffer(days);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        Assert.Equal(expected, dt);
        Assert.Equal(TimeSpan.Zero, dt.TimeOfDay);
    }

    [Fact]
    public void Decode_Date64Array_YieldsDateTime()
    {
        // date64 is rare (Java tooling produces it), but external writers
        // may surface it; verify the decoder handles it.
        var expected = new DateTime(2024, 1, 15, 0, 0, 0, DateTimeKind.Unspecified);
        var ms = (long)(expected - new DateTime(1970, 1, 1)).TotalMilliseconds;
        var buf = BuildDate64ScalarBuffer(ms);

        var decoded = ArrowMarshal.Decode(buf);

        var dt = Assert.IsType<DateTime>(decoded);
        Assert.Equal(expected, dt);
    }

    [Fact]
    public void Decode_TimestampVector_YieldsDateTimeArray()
    {
        // Three timestamps in a vector — the path the kernel uses to ship a
        // 1-D pandas datetime series back to Excel.
        var dates = new[]
        {
            new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Unspecified),
            new DateTime(2024, 6, 15, 12, 30, 0, DateTimeKind.Unspecified),
            new DateTime(2024, 12, 31, 23, 59, 59, DateTimeKind.Unspecified),
        };
        var epoch = new DateTime(1970, 1, 1).Ticks;
        var micros = new long[dates.Length];
        for (var i = 0; i < dates.Length; i++)
            micros[i] = (dates[i].Ticks - epoch) / 10L;
        var buf = BuildTimestampVectorBuffer(micros, TimeUnit.Microsecond);

        var decoded = ArrowMarshal.Decode(buf);

        var vec = Assert.IsType<object?[]>(decoded);
        Assert.Equal(3, vec.Length);
        for (var i = 0; i < dates.Length; i++)
            Assert.Equal(dates[i], (DateTime)vec[i]!);
    }

    // -------------------------------------------------------------------------
    // Embedded nulls — verify nulls survive the round-trip in every
    // primitive column type. The all-null and "some nulls in a string
    // column" paths are covered above; these pin the numeric / bool paths.
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeTable_NullInNumericColumn_PreservesNull()
    {
        var input = new object?[,]
        {
            { 1.0, 10.0 },
            { null, 20.0 },
            { 3.0, null },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_NullInBoolColumn_PreservesNull()
    {
        var input = new object?[,]
        {
            { true, false },
            { null, true },
            { false, null },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeVector_NullInNumericVector_PreservesNull()
    {
        var input = new object?[] { 1.0, null, 3.0, null, 5.0 };
        var decoded = (object?[])ArrowMarshal.Decode(ArrowMarshal.EncodeVector(input))!;
        Assert.Equal(input, decoded);
    }

    // -------------------------------------------------------------------------
    // Formula-result regression — Excel's Value2 unwraps a formula cell to
    // its calculated value (double / string / bool). These tests pin that
    // the marshalling layer carries that through unchanged, since changes
    // to the type-inference table have historically broken this path.
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeTable_FormulaResultDouble_RoundTripsAsDouble()
    {
        // A column where every cell is the result of =SUM(...) → all
        // doubles. Should land on the numeric fast path.
        var input = new object?[,]
        {
            { 3.0 },
            { 7.5 },
            { -2.25 },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_FormulaResultString_RoundTripsAsString()
    {
        // =TEXT(...) returns strings; verify the column infers string
        // rather than tripping over a mixed-type fallback.
        var input = new object?[,]
        {
            { "2024-01-15" },
            { "2024-02-20" },
            { "2024-03-25" },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        AssertTablesEqual(input, decoded);
    }

    [Fact]
    public void EncodeTable_FormulaMixedNumericAndError_FallsBackToString()
    {
        // A formula column where some cells errored (#DIV/0!) — Excel
        // surfaces those as the string "#DIV/0!" via Value2 (with
        // AllowReference=false on the UDF), mixing strings into a numeric
        // column. The column should fall back to string, not throw, and
        // each value should round-trip via ToString().
        var input = new object?[,]
        {
            { 1.0 },
            { "#DIV/0!" },
            { 3.0 },
        };
        var decoded = (object?[,])ArrowMarshal.Decode(ArrowMarshal.EncodeTable(input))!;
        // Numbers go through ToString() on the string fallback path; that's
        // "1" / "3" via the invariant-culture default for double.
        Assert.Equal("1", decoded[0, 0]);
        Assert.Equal("#DIV/0!", decoded[1, 0]);
        Assert.Equal("3", decoded[2, 0]);
    }

    // -------------------------------------------------------------------------
    // Argument validation
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeTable_NullValues_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ArrowMarshal.EncodeTable(null!));
    }

    [Fact]
    public void EncodeVector_NullValues_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ArrowMarshal.EncodeVector(null!));
    }

    [Fact]
    public void Decode_NullBuffer_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ArrowMarshal.Decode(null!));
    }

    [Fact]
    public void PeekShape_NullBuffer_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ArrowMarshal.PeekShape(null!));
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static void AssertTablesEqual(object?[,] expected, object?[,] actual)
    {
        Assert.Equal(expected.GetLength(0), actual.GetLength(0));
        Assert.Equal(expected.GetLength(1), actual.GetLength(1));
        for (var r = 0; r < expected.GetLength(0); r++)
            for (var c = 0; c < expected.GetLength(1); c++)
                Assert.Equal(expected[r, c], actual[r, c]);
    }

    /// <summary>
    /// Write a one-column, one-row Arrow IPC stream with the supplied
    /// <paramref name="array"/> and the PyExcel shape/orientation metadata.
    /// Used by the date / timestamp decode tests, where the array types
    /// aren't produced by <see cref="ArrowMarshal"/>'s own encoders.
    /// </summary>
    private static byte[] BuildSingleColumnBuffer(
        IArrowType columnType,
        IArrowArray array,
        int length,
        string shape,
        string? orientation = null)
    {
        var field = new Field("0", columnType, nullable: true);
        var metadata = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["pyexcel-shape"] = shape,
        };
        if (orientation is { }) metadata["pyexcel-orientation"] = orientation;
        var schema = new Schema.Builder().Field(field).Metadata(metadata).Build();

        var batch = new RecordBatch(schema, new[] { array }, length);
        using var ms = new MemoryStream();
        using (var writer = new ArrowStreamWriter(ms, schema))
        {
            writer.WriteRecordBatch(batch);
            writer.WriteEnd();
        }
        return ms.ToArray();
    }

    private static byte[] BuildTimestampScalarBuffer(long value, TimeUnit unit)
    {
        // Apache.Arrow's TimestampType takes a `string timezone` (default "")
        // — empty signals "naive", matching pyarrow's default.
        var type = new TimestampType(unit, timezone: "");
        var builder = new TimestampArray.Builder(type);
        builder.Append(value);
        return BuildSingleColumnBuffer(type, builder.Build(), length: 1, shape: "scalar");
    }

    private static byte[] BuildTimestampVectorBuffer(long[] values, TimeUnit unit)
    {
        var type = new TimestampType(unit, timezone: "");
        var builder = new TimestampArray.Builder(type);
        foreach (var v in values) builder.Append(v);
        return BuildSingleColumnBuffer(
            type, builder.Build(), length: values.Length,
            shape: "vector", orientation: "column");
    }

    private static byte[] BuildDate32ScalarBuffer(int daysSinceEpoch)
    {
        var builder = new Date32Array.Builder();
        builder.Append(daysSinceEpoch);
        return BuildSingleColumnBuffer(
            Date32Type.Default, builder.Build(), length: 1, shape: "scalar");
    }

    private static byte[] BuildDate64ScalarBuffer(long msSinceEpoch)
    {
        var builder = new Date64Array.Builder();
        builder.Append(msSinceEpoch);
        return BuildSingleColumnBuffer(
            Date64Type.Default, builder.Build(), length: 1, shape: "scalar");
    }
}
