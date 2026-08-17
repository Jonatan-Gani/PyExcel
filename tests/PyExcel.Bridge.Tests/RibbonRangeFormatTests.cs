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
    [InlineData("prices:dataframe=Sheet1!A1:C10")]
    [InlineData(":list=A1:A10")]
    [InlineData("a:dict=A1:B9; B1; c:scalar=Sheet2!C3; d=D1:D4")]
    public void Parse_Then_Format_RoundTrips(string input)
    {
        var bindings = RibbonRangeParser.Parse(input);
        var reparsed = RibbonRangeParser.Parse(RibbonRangeParser.Format(bindings));
        Assert.Equal(bindings.Count, reparsed.Count);
        for (int i = 0; i < bindings.Count; i++)
        {
            // Compare the records whole. Field-by-field assertions let a new
            // member (the declared type was exactly this case) be dropped by
            // Format while the round-trip test stayed green.
            Assert.Equal(bindings[i], reparsed[i]);
        }
    }

    [Fact]
    public void Format_Null_Throws()
        => Assert.Throws<System.ArgumentNullException>(() => RibbonRangeParser.Format(null!));
}
