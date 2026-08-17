using System;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// The <c>name:type=range</c> half of the typed I/O contract — the vocabulary
/// itself, and the grammar extension that carries a declared type through the
/// binding text. See <c>docs/typed-io-contract.md</c>.
/// </summary>
public class DeclaredTypeGrammarTests
{
    // -------------------------------------------------------------------------
    // Vocabulary
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(PyExcelType.Auto, "auto")]
    [InlineData(PyExcelType.DataFrame, "dataframe")]
    [InlineData(PyExcelType.Series, "series")]
    [InlineData(PyExcelType.List, "list")]
    [InlineData(PyExcelType.Tuple, "tuple")]
    [InlineData(PyExcelType.Set, "set")]
    [InlineData(PyExcelType.Dict, "dict")]
    [InlineData(PyExcelType.NDArray, "ndarray")]
    [InlineData(PyExcelType.Scalar, "scalar")]
    public void WireName_MatchesThePythonMirror(PyExcelType type, string expected)
        => Assert.Equal(expected, PyExcelTypes.WireName(type));

    [Fact]
    public void EveryTypeHasAWireNameAndADisplayName()
    {
        foreach (var type in PyExcelTypes.All)
        {
            Assert.False(string.IsNullOrWhiteSpace(PyExcelTypes.WireName(type)));
            Assert.False(string.IsNullOrWhiteSpace(PyExcelTypes.DisplayName(type)));
        }
    }

    [Theory]
    [InlineData("dataframe", PyExcelType.DataFrame)]
    [InlineData("DataFrame", PyExcelType.DataFrame)]
    [InlineData("  NDARRAY  ", PyExcelType.NDArray)]
    public void TryParse_IsCaseAndWhitespaceInsensitive(string token, PyExcelType expected)
    {
        Assert.True(PyExcelTypes.TryParse(token, out var parsed));
        Assert.Equal(expected, parsed);
    }

    [Theory]
    [InlineData("")]
    [InlineData(null)]
    [InlineData("frobnicate")]
    [InlineData("data frame")]
    public void TryParse_RejectsAnythingUnrecognised(string? token)
        => Assert.False(PyExcelTypes.TryParse(token, out _));

    [Theory]
    [InlineData(1, 1, PyExcelType.Scalar)]
    [InlineData(1, 8, PyExcelType.List)]
    [InlineData(8, 1, PyExcelType.List)]
    [InlineData(8, 3, PyExcelType.DataFrame)]
    public void ResolveAuto_FollowsTheDocumentedDefaults(
        int rows, int cols, PyExcelType expected)
        => Assert.Equal(expected, PyExcelTypes.ResolveAuto(PyExcelType.Auto, rows, cols));

    [Fact]
    public void ResolveAuto_LeavesADeclaredTypeUntouched()
        => Assert.Equal(
            PyExcelType.Set, PyExcelTypes.ResolveAuto(PyExcelType.Set, 1, 1));

    // -------------------------------------------------------------------------
    // Grammar
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_ReadsANamedTypedBinding()
    {
        var b = Assert.Single(RibbonRangeParser.Parse("prices:dataframe=Sheet1!A1:C10"));
        Assert.Equal("prices", b.Name);
        Assert.Equal("Sheet1!A1:C10", b.RangeText);
        Assert.Equal(PyExcelType.DataFrame, b.DeclaredType);
    }

    [Fact]
    public void Parse_ReadsAnAnonymousTypedBinding()
    {
        var b = Assert.Single(RibbonRangeParser.Parse(":list=A1:A10"));
        Assert.Null(b.Name);
        Assert.Equal("A1:A10", b.RangeText);
        Assert.Equal(PyExcelType.List, b.DeclaredType);
    }

    [Fact]
    public void Parse_DefaultsToAutoWhenNoTypeIsDeclared()
    {
        var b = Assert.Single(RibbonRangeParser.Parse("prices=A1:C10"));
        Assert.Equal(PyExcelType.Auto, b.DeclaredType);
    }

    [Fact]
    public void Parse_LeavesAnExistingNameContainingAColonIntact()
    {
        // Backward compatibility: the split only counts when the text after
        // the final colon is a known type, so a saved name keeps its colon
        // rather than being silently reinterpreted.
        var b = Assert.Single(RibbonRangeParser.Parse("my:name=A1"));
        Assert.Equal("my:name", b.Name);
        Assert.Equal(PyExcelType.Auto, b.DeclaredType);
    }

    [Fact]
    public void Parse_SplitsOnTheFinalColonOnly()
    {
        var b = Assert.Single(RibbonRangeParser.Parse("a:b:scalar=A1"));
        Assert.Equal("a:b", b.Name);
        Assert.Equal(PyExcelType.Scalar, b.DeclaredType);
    }

    [Fact]
    public void Parse_DoesNotReadATypeOutOfAnUnnamedRange()
    {
        // 'A1:C10' has a colon but no '=', so nothing is a name or a type.
        var b = Assert.Single(RibbonRangeParser.Parse("A1:C10"));
        Assert.Null(b.Name);
        Assert.Equal("A1:C10", b.RangeText);
        Assert.Equal(PyExcelType.Auto, b.DeclaredType);
    }

    [Fact]
    public void Parse_StillRejectsAnEmptyNameWithNoType()
        => Assert.Throws<FormatException>(() => RibbonRangeParser.Parse("=A1:C10"));

    [Fact]
    public void Parse_TypeTokenIsCaseInsensitive()
    {
        var b = Assert.Single(RibbonRangeParser.Parse("x:DataFrame=A1:C3"));
        Assert.Equal(PyExcelType.DataFrame, b.DeclaredType);
    }

    [Theory]
    [InlineData(PyExcelType.Auto, "x=A1")]
    [InlineData(PyExcelType.DataFrame, "x:dataframe=A1")]
    [InlineData(PyExcelType.Set, "x:set=A1")]
    public void Format_EmitsTheTypeOnlyWhenDeclared(PyExcelType type, string expected)
        => Assert.Equal(
            expected,
            RibbonRangeParser.Format(new[] { new RangeBinding("x", "A1", type) }));

    [Fact]
    public void Format_EmitsAnAnonymousTypedBindingWithALeadingColon()
        => Assert.Equal(
            ":scalar=A1",
            RibbonRangeParser.Format(
                new[] { new RangeBinding(null, "A1", PyExcelType.Scalar) }));

    [Fact]
    public void MixedBindings_RoundTripEveryField()
    {
        const string text = "sales:dataframe=Sheet1!A1:C10; E1; rate:scalar=Sheet1!E1";
        var parsed = RibbonRangeParser.Parse(text);

        Assert.Equal(3, parsed.Count);
        Assert.Equal(PyExcelType.DataFrame, parsed[0].DeclaredType);
        Assert.Equal(PyExcelType.Auto, parsed[1].DeclaredType);
        Assert.Equal(PyExcelType.Scalar, parsed[2].DeclaredType);
        Assert.Equal(text, RibbonRangeParser.Format(parsed));
    }
}
