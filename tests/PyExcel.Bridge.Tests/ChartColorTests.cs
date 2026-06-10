using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="ChartColor"/> turns the colour strings a chart spec carries
/// (hex, rgb()/rgba(), CSS names — the union of what Plotly emits and what
/// v1's <c>chartBuilder.bas</c> accepted) into Excel's OLE BGR-packed
/// integer. Unrecognised input is black, never an exception — a wrong
/// colour must not cost the user the chart.
/// </summary>
public class ChartColorTests
{
    private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

    [Theory]
    [InlineData("#ff0000", 255, 0, 0)]
    [InlineData("#00ff00", 0, 255, 0)]
    [InlineData("#0000ff", 0, 0, 255)]
    [InlineData("#123456", 0x12, 0x34, 0x56)]
    [InlineData("#FFFFFF", 255, 255, 255)]
    public void Parse_SixDigitHex(string spec, int r, int g, int b)
    {
        Assert.Equal(Rgb(r, g, b), ChartColor.Parse(spec));
    }

    [Fact]
    public void Parse_ThreeDigitHex_Expands()
    {
        Assert.Equal(Rgb(0xff, 0x00, 0xaa), ChartColor.Parse("#f0a"));
    }

    [Fact]
    public void Parse_RgbFunction()
    {
        Assert.Equal(Rgb(10, 20, 30), ChartColor.Parse("rgb(10, 20, 30)"));
    }

    [Fact]
    public void Parse_RgbaFunction_IgnoresAlpha()
    {
        // Plotly themes routinely emit rgba(); alpha rides separately as
        // fill opacity, so only the channels matter here.
        Assert.Equal(Rgb(99, 110, 250), ChartColor.Parse("rgba(99, 110, 250, 0.5)"));
    }

    [Fact]
    public void Parse_RgbWithFloatChannels_Rounds()
    {
        Assert.Equal(Rgb(11, 20, 30), ChartColor.Parse("rgb(10.5, 20.4, 29.6)"));
    }

    [Fact]
    public void Parse_RgbChannelsOutOfRange_Clamped()
    {
        Assert.Equal(Rgb(255, 0, 255), ChartColor.Parse("rgb(300, -5, 255)"));
    }

    [Theory]
    [InlineData("red", 255, 0, 0)]
    [InlineData("green", 0, 128, 0)]
    [InlineData("blue", 0, 0, 255)]
    [InlineData("Gray", 128, 128, 128)]
    [InlineData("GREY", 128, 128, 128)]
    [InlineData("orange", 255, 165, 0)]
    [InlineData("white", 255, 255, 255)]
    public void Parse_NamedColors(string spec, int r, int g, int b)
    {
        Assert.Equal(Rgb(r, g, b), ChartColor.Parse(spec));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("#12")]          // malformed hex
    [InlineData("#zzzzzz")]      // non-hex digits
    [InlineData("rgb(1,2)")]     // wrong arity
    [InlineData("chartreuse")]   // unmapped name
    public void Parse_UnrecognisedInput_IsBlack(string? spec)
    {
        Assert.Equal(0, ChartColor.Parse(spec));
    }
}
