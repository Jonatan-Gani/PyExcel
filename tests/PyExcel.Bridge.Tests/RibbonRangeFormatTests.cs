using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers <see cref="RibbonRangeParser.Format"/> — the inverse the Note 4
/// list editor uses to serialise its rows back to the ribbon's
/// <c>{name}=range; …</c> syntax — and its round-trip with <c>Parse</c>.
/// </summary>
public class RibbonRangeFormatTests
{
    [Fact]
    public void Format_Empty_ReturnsEmptyString()
        => Assert.Equal("", RibbonRangeParser.Format(new RangeBinding[0]));

    [Fact]
    public void Format_AnonymousRange_IsJustTheRange()
        => Assert.Equal("A1:B2",
            RibbonRangeParser.Format(new[] { new RangeBinding(null, "A1:B2") }));

    [Fact]
    public void Format_NamedRange_UsesNameEquals()
        => Assert.Equal("prices=A1:B2",
            RibbonRangeParser.Format(new[] { new RangeBinding("prices", "A1:B2") }));

    [Fact]
    public void Format_Mixed_JoinsWithSemicolonSpace()
    {
        var text = RibbonRangeParser.Format(new[]
        {
            new RangeBinding("prices", "Sheet1!A1:C10"),
            new RangeBinding(null, "E1"),
        });
        Assert.Equal("prices=Sheet1!A1:C10; E1", text);
    }

    [Fact]
    public void Format_SkipsBlankRangeText()
        => Assert.Equal("A1", RibbonRangeParser.Format(new[]
        {
            new RangeBinding(null, "A1"),
            new RangeBinding("x", "   "),
        }));

    [Fact]
    public void Format_TrimsNameAndRange()
        => Assert.Equal("x=A1",
            RibbonRangeParser.Format(new[] { new RangeBinding("  x ", "  A1 ") }));

    [Theory]
    [InlineData("A1:C10")]
    [InlineData("prices=A1:C10")]
    [InlineData("prices=Sheet1!A1:C10; E1; tax=Sheet2!B2")]
    public void Parse_Then_Format_RoundTrips(string input)
    {
        var bindings = RibbonRangeParser.Parse(input);
        var reparsed = RibbonRangeParser.Parse(RibbonRangeParser.Format(bindings));
        Assert.Equal(bindings.Count, reparsed.Count);
        for (int i = 0; i < bindings.Count; i++)
        {
            Assert.Equal(bindings[i].Name, reparsed[i].Name);
            Assert.Equal(bindings[i].RangeText, reparsed[i].RangeText);
        }
    }

    [Fact]
    public void Format_Null_Throws()
        => Assert.Throws<System.ArgumentNullException>(() => RibbonRangeParser.Format(null!));
}
