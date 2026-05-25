using System;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class RibbonRangeParserTests
{
    // -------------------------------------------------------------------------
    // Empty input
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\t \n")]
    public void Parse_NullOrWhitespace_ReturnsEmpty(string? input)
    {
        Assert.Empty(RibbonRangeParser.Parse(input));
    }

    // -------------------------------------------------------------------------
    // Anonymous bindings
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_SingleAnonymousRange_OneBindingNullName()
    {
        var result = RibbonRangeParser.Parse("A1:C10");
        var b = Assert.Single(result);
        Assert.Null(b.Name);
        Assert.Equal("A1:C10", b.RangeText);
    }

    [Fact]
    public void Parse_TrimsSurroundingWhitespace()
    {
        var result = RibbonRangeParser.Parse("   A1:C10   ");
        var b = Assert.Single(result);
        Assert.Equal("A1:C10", b.RangeText);
    }

    [Fact]
    public void Parse_SheetQualifiedRange_PreservedVerbatim()
    {
        var result = RibbonRangeParser.Parse("Sheet1!A1:C10");
        var b = Assert.Single(result);
        Assert.Equal("Sheet1!A1:C10", b.RangeText);
    }

    // -------------------------------------------------------------------------
    // Named bindings
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_SingleNamedBinding()
    {
        var result = RibbonRangeParser.Parse("prices=A1:C10");
        var b = Assert.Single(result);
        Assert.Equal("prices", b.Name);
        Assert.Equal("A1:C10", b.RangeText);
    }

    [Fact]
    public void Parse_NamedBinding_TrimsAroundEquals()
    {
        var result = RibbonRangeParser.Parse("  prices  =  A1:C10  ");
        var b = Assert.Single(result);
        Assert.Equal("prices", b.Name);
        Assert.Equal("A1:C10", b.RangeText);
    }

    // -------------------------------------------------------------------------
    // Multiple bindings — ordering, separators
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_MultipleNamedBindings_PreservesOrder()
    {
        var result = RibbonRangeParser.Parse("prices=A1:C10; signals=D1:D10");
        Assert.Equal(2, result.Count);
        Assert.Equal("prices", result[0].Name);
        Assert.Equal("A1:C10", result[0].RangeText);
        Assert.Equal("signals", result[1].Name);
        Assert.Equal("D1:D10", result[1].RangeText);
    }

    [Fact]
    public void Parse_MixedNamedAndAnonymous_AllowedInOrder()
    {
        var result = RibbonRangeParser.Parse("A1:C10; signals=D1:D10");
        Assert.Equal(2, result.Count);
        Assert.Null(result[0].Name);
        Assert.Equal("A1:C10", result[0].RangeText);
        Assert.Equal("signals", result[1].Name);
        Assert.Equal("D1:D10", result[1].RangeText);
    }

    [Fact]
    public void Parse_TrailingSemicolon_Ignored()
    {
        var result = RibbonRangeParser.Parse("prices=A1:C10;");
        Assert.Single(result);
    }

    [Fact]
    public void Parse_DoubleSemicolon_Ignored()
    {
        var result = RibbonRangeParser.Parse("prices=A1:C10;;signals=D1:D10");
        Assert.Equal(2, result.Count);
    }

    [Fact]
    public void Parse_LeadingSemicolon_Ignored()
    {
        var result = RibbonRangeParser.Parse(";prices=A1:C10");
        Assert.Single(result);
    }

    // -------------------------------------------------------------------------
    // Malformed input
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_EmptyNameBeforeEquals_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => RibbonRangeParser.Parse("=A1:C10"));
        Assert.Contains("empty name", ex.Message);
    }

    [Fact]
    public void Parse_EmptyRangeAfterEquals_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => RibbonRangeParser.Parse("prices="));
        Assert.Contains("empty range", ex.Message);
    }

    [Fact]
    public void Parse_DuplicateNames_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => RibbonRangeParser.Parse("prices=A1:A5; prices=B1:B5"));
        Assert.Contains("duplicate", ex.Message);
    }

    [Fact]
    public void Parse_DuplicateNames_CaseSensitive_DistinctNamesAllowed()
    {
        // Ribbon names are case-sensitive (matches StateService keying
        // throughout the codebase).
        var result = RibbonRangeParser.Parse("Prices=A1:A5; prices=B1:B5");
        Assert.Equal(2, result.Count);
    }
}
