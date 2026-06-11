using System.Collections.Generic;
using PyExcel.Forms;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class KwargsTextTests
{
    // -------------------------------------------------------------------------
    // TryParse — happy paths
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("\n\n  \n")]
    public void TryParse_BlankInput_EmptyDictionaryNoError(string? text)
    {
        var result = KwargsText.TryParse(text, out var error);
        Assert.Null(error);
        Assert.NotNull(result);
        Assert.Empty(result!);
    }

    [Fact]
    public void TryParse_SinglePair_Parsed()
    {
        var result = KwargsText.TryParse("alpha=1", out var error);
        Assert.Null(error);
        Assert.Single(result!);
        Assert.Equal("1", result!["alpha"]);
    }

    [Fact]
    public void TryParse_MultiplePairs_AllParsed()
    {
        var result = KwargsText.TryParse("alpha=1\nbeta=two\ngamma=3.5", out var error);
        Assert.Null(error);
        Assert.Equal(3, result!.Count);
        Assert.Equal("1", result["alpha"]);
        Assert.Equal("two", result["beta"]);
        Assert.Equal("3.5", result["gamma"]);
    }

    [Theory]
    [InlineData("a=1\r\nb=2")]   // CRLF
    [InlineData("a=1\rb=2")]     // bare CR
    [InlineData("a=1\nb=2")]     // LF
    public void TryParse_AllLineEndings_Tolerated(string text)
    {
        var result = KwargsText.TryParse(text, out var error);
        Assert.Null(error);
        Assert.Equal(2, result!.Count);
        Assert.Equal("1", result["a"]);
        Assert.Equal("2", result["b"]);
    }

    [Fact]
    public void TryParse_BlankLinesBetweenPairs_Ignored()
    {
        var result = KwargsText.TryParse("a=1\n\n   \nb=2\n", out var error);
        Assert.Null(error);
        Assert.Equal(2, result!.Count);
    }

    [Fact]
    public void TryParse_ValueWithEquals_SplitsOnFirstOnly()
    {
        var result = KwargsText.TryParse("expr=a==b", out var error);
        Assert.Null(error);
        Assert.Equal("a==b", result!["expr"]);
    }

    [Fact]
    public void TryParse_KeyAndValueTrimmed()
    {
        var result = KwargsText.TryParse("  key   =   value here  ", out var error);
        Assert.Null(error);
        Assert.Equal("value here", result!["key"]);
        Assert.True(result.ContainsKey("key"));
    }

    [Fact]
    public void TryParse_EmptyValueAllowed()
    {
        var result = KwargsText.TryParse("flag=", out var error);
        Assert.Null(error);
        Assert.Equal(string.Empty, result!["flag"]);
    }

    // -------------------------------------------------------------------------
    // TryParse — error paths
    // -------------------------------------------------------------------------

    [Fact]
    public void TryParse_LineMissingEquals_Fails()
    {
        var result = KwargsText.TryParse("alpha=1\nnoequalshere", out var error);
        Assert.Null(result);
        Assert.NotNull(error);
        Assert.Contains("noequalshere", error!);
    }

    [Fact]
    public void TryParse_BlankKey_Fails()
    {
        var result = KwargsText.TryParse("=value", out var error);
        Assert.Null(result);
        Assert.NotNull(error);
    }

    [Fact]
    public void TryParse_DuplicateKey_Fails()
    {
        var result = KwargsText.TryParse("a=1\na=2", out var error);
        Assert.Null(result);
        Assert.NotNull(error);
        Assert.Contains("a", error!);
    }

    [Fact]
    public void TryParse_KeysAreCaseSensitive()
    {
        var result = KwargsText.TryParse("a=1\nA=2", out var error);
        Assert.Null(error);
        Assert.Equal(2, result!.Count);
    }

    // -------------------------------------------------------------------------
    // Format
    // -------------------------------------------------------------------------

    [Fact]
    public void Format_Null_Empty()
    {
        Assert.Equal(string.Empty, KwargsText.Format(null));
    }

    [Fact]
    public void Format_EmptyDictionary_Empty()
    {
        Assert.Equal(string.Empty, KwargsText.Format(new Dictionary<string, string>()));
    }

    [Fact]
    public void Format_Pairs_OnePerLine()
    {
        var text = KwargsText.Format(new Dictionary<string, string>
        {
            ["a"] = "1",
            ["b"] = "two",
        });
        Assert.Equal("a=1\nb=two", text);
    }

    // -------------------------------------------------------------------------
    // Round-trip
    // -------------------------------------------------------------------------

    [Fact]
    public void RoundTrip_FormatThenParse_PreservesPairs()
    {
        var original = new Dictionary<string, string>
        {
            ["alpha"] = "1",
            ["beta"] = "a==b",
            ["empty"] = "",
        };
        var text = KwargsText.Format(original);
        var parsed = KwargsText.TryParse(text, out var error);
        Assert.Null(error);
        Assert.Equal(original.Count, parsed!.Count);
        foreach (var kv in original)
            Assert.Equal(kv.Value, parsed[kv.Key]);
    }
}
