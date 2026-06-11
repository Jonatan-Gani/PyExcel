using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ProgressModelTests
{
    [Theory]
    [InlineData(0, 0)]
    [InlineData(50, 50)]
    [InlineData(100, 100)]
    [InlineData(-5, 0)]
    [InlineData(150, 100)]
    [InlineData(42.6, 43)]
    public void ClampPercent_ClampsAndRounds(double input, int expected)
    {
        Assert.Equal(expected, ProgressModel.ClampPercent(input));
    }

    [Fact]
    public void ClampPercent_NaN_Zero()
    {
        Assert.Equal(0, ProgressModel.ClampPercent(double.NaN));
    }

    [Fact]
    public void FormatLine_PercentAndMessage()
    {
        Assert.Equal("42% — crunching", ProgressModel.FormatLine(42, "crunching"));
    }

    [Fact]
    public void FormatLine_PercentOnly()
    {
        Assert.Equal("42%", ProgressModel.FormatLine(42, ""));
    }

    [Fact]
    public void FormatLine_IndeterminateWithMessage()
    {
        Assert.Equal("loading data", ProgressModel.FormatLine(null, "loading data"));
    }

    [Fact]
    public void FormatLine_IndeterminateNoMessage_Default()
    {
        Assert.Equal("Working…", ProgressModel.FormatLine(null, null));
    }

    [Fact]
    public void FormatLine_ClampsPercentInLine()
    {
        Assert.Equal("100% — done", ProgressModel.FormatLine(250, "done"));
    }
}
