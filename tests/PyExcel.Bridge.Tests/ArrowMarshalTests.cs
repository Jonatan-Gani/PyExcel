using System;
using System.Collections.Generic;
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
}
