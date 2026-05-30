using System;
using System.Globalization;
using System.Threading;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class CsvCellFormatterTests
{
    [Fact]
    public void Format_Null_ReturnsNull()
    {
        Assert.Null(CsvCellFormatter.Format(null));
    }

    [Fact]
    public void Format_String_PassThrough()
    {
        Assert.Equal("hello", CsvCellFormatter.Format("hello"));
    }

    [Fact]
    public void Format_EmptyString_ReturnsEmptyString()
    {
        // Distinct from null — an empty string is a known explicit value.
        Assert.Equal("", CsvCellFormatter.Format(""));
    }

    // -------------------------------------------------------------------------
    // Numerics — invariant round-trip format
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(0.0, "0")]
    [InlineData(1.0, "1")]
    [InlineData(-1.0, "-1")]
    [InlineData(3.14, "3.14")]
    public void Format_Double_UsesInvariantRoundTrip(double value, string expected)
    {
        Assert.Equal(expected, CsvCellFormatter.Format(value));
    }

    [Fact]
    public void Format_DoubleUnderGermanCulture_StillEmitsDotDecimal()
    {
        var saved = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal("3.14", CsvCellFormatter.Format(3.14));
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = saved;
        }
    }

    // -------------------------------------------------------------------------
    // Booleans — TRUE/FALSE, matches inference round-trip
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(true, "TRUE")]
    [InlineData(false, "FALSE")]
    public void Format_Bool_TrueFalseUppercase(bool value, string expected)
    {
        Assert.Equal(expected, CsvCellFormatter.Format(value));
    }

    [Fact]
    public void Format_Bool_RoundTripsThroughInference()
    {
        // Inference + Format must be symmetric for the supported types.
        var formatted = CsvCellFormatter.Format(true);
        Assert.Equal(true, CsvCellTypeInference.Infer(formatted));
    }

    // -------------------------------------------------------------------------
    // DateTime — ISO 8601
    // -------------------------------------------------------------------------

    [Fact]
    public void Format_DateTime_UsesIso8601()
    {
        var dt = new DateTime(2026, 5, 30, 14, 30, 45);
        Assert.Equal("2026-05-30T14:30:45", CsvCellFormatter.Format(dt));
    }

    [Fact]
    public void Format_DateTimeUnderGermanCulture_StillIso8601()
    {
        var saved = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            var dt = new DateTime(2026, 5, 30);
            Assert.Equal("2026-05-30T00:00:00", CsvCellFormatter.Format(dt));
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = saved;
        }
    }

    // -------------------------------------------------------------------------
    // Other types — invariant ToString fallback
    // -------------------------------------------------------------------------

    [Fact]
    public void Format_Int_UsesInvariantToString()
    {
        Assert.Equal("42", CsvCellFormatter.Format(42));
    }

    [Fact]
    public void Format_Decimal_UsesInvariantToString()
    {
        // Decimal is the prototypical "other type" — Range.Value2 won't
        // emit them, but a kernel return that round-trips through the
        // formatter shouldn't depend on the host's culture.
        var saved = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal("3.14", CsvCellFormatter.Format(3.14m));
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = saved;
        }
    }
}
