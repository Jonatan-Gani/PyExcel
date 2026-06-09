using System;
using System.Globalization;
using System.Threading;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class CsvCellTypeInferenceTests
{
    // -------------------------------------------------------------------------
    // Null / empty
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Infer_NullOrEmpty_ReturnsNull(string? input)
    {
        Assert.Null(CsvCellTypeInference.Infer(input));
    }

    // -------------------------------------------------------------------------
    // Booleans
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("TRUE", true)]
    [InlineData("true", true)]
    [InlineData("True", true)]
    [InlineData("FALSE", false)]
    [InlineData("false", false)]
    [InlineData("False", false)]
    public void Infer_BoolTokens_ReturnBool(string input, bool expected)
    {
        Assert.Equal(expected, CsvCellTypeInference.Infer(input));
    }

    [Theory]
    [InlineData("yes")]
    [InlineData("no")]
    [InlineData("T")]
    [InlineData("F")]
    public void Infer_NonStandardBoolTokens_ReturnString(string input)
    {
        Assert.Equal(input, CsvCellTypeInference.Infer(input));
    }

    // -------------------------------------------------------------------------
    // Numbers
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("0", 0.0)]
    [InlineData("1", 1.0)]
    [InlineData("-1", -1.0)]
    [InlineData("3.14", 3.14)]
    [InlineData("-3.14", -3.14)]
    [InlineData("1.5e2", 150.0)]
    [InlineData("0.5", 0.5)]
    public void Infer_Numbers_ReturnDouble(string input, double expected)
    {
        Assert.Equal(expected, CsvCellTypeInference.Infer(input));
    }

    [Fact]
    public void Infer_NumberWithThousandsSeparator_Parses()
    {
        Assert.Equal(1000.0, CsvCellTypeInference.Infer("1,000"));
    }

    // -------------------------------------------------------------------------
    // Leading-zero / leading-plus guards
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("00")]
    [InlineData("01")]
    [InlineData("0123")]
    [InlineData("01.5")]
    public void Infer_LeadingZero_StaysAsString(string input)
    {
        // A user with a column of SKUs starting with zero would lose the
        // zero on round-trip if we parsed these as numbers — keep as
        // string.
        Assert.Equal(input, CsvCellTypeInference.Infer(input));
    }

    [Theory]
    [InlineData("+1")]
    [InlineData("+44")]
    [InlineData("+1.5")]
    public void Infer_LeadingPlus_StaysAsString(string input)
    {
        // Phone-number prefixes etc. — Excel keeps these as text.
        Assert.Equal(input, CsvCellTypeInference.Infer(input));
    }

    // -------------------------------------------------------------------------
    // Locale independence
    // -------------------------------------------------------------------------

    [Fact]
    public void Infer_NumberUnderGermanCulture_StillParsesAsInvariant()
    {
        // The CSV format uses dot for decimals; we must parse under
        // invariant culture so a German-locale Excel doesn't change
        // behaviour. Restore the original culture in finally so other
        // tests aren't perturbed.
        var saved = Thread.CurrentThread.CurrentCulture;
        try
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
            Assert.Equal(3.14, CsvCellTypeInference.Infer("3.14"));
        }
        finally
        {
            Thread.CurrentThread.CurrentCulture = saved;
        }
    }

    // -------------------------------------------------------------------------
    // Plain strings
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("hello")]
    [InlineData("hello world")]
    [InlineData("123abc")]
    [InlineData("3.14.15")]
    public void Infer_PlainStrings_PassThrough(string input)
    {
        Assert.Equal(input, CsvCellTypeInference.Infer(input));
    }
}
