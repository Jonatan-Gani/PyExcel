using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class LegacyFormulaDecoderTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Decode_BlankInput_ReturnsNull(string? input)
    {
        Assert.Null(LegacyFormulaDecoder.Decode(input));
    }

    [Fact]
    public void Decode_SimpleLiteral()
    {
        Assert.Equal("A1:C10", LegacyFormulaDecoder.Decode("=\"A1:C10\""));
    }

    [Fact]
    public void Decode_LiteralWithoutLeadingEquals()
    {
        // Some readers hand back the formula without the '='.
        Assert.Equal("F1", LegacyFormulaDecoder.Decode("\"F1\""));
    }

    [Fact]
    public void Decode_ConcatenatedChunks()
    {
        // v1 split values over 255 chars into "a"&"b"&"c".
        Assert.Equal("abc", LegacyFormulaDecoder.Decode("=\"a\"&\"b\"&\"c\""));
    }

    [Fact]
    public void Decode_ConcatenatedWithWhitespaceAroundAmpersand()
    {
        Assert.Equal("onetwo", LegacyFormulaDecoder.Decode("=\"one\" & \"two\""));
    }

    [Fact]
    public void Decode_EscapedDoubleQuotes()
    {
        // "" inside a literal is a single embedded quote.
        Assert.Equal("a\"b", LegacyFormulaDecoder.Decode("=\"a\"\"b\""));
    }

    [Fact]
    public void Decode_EscapedQuotesAcrossChunks()
    {
        Assert.Equal("x\"y", LegacyFormulaDecoder.Decode("=\"x\"\"\"&\"y\""));
    }

    [Fact]
    public void Decode_EmptyLiteral_ReturnsEmptyString()
    {
        Assert.Equal(string.Empty, LegacyFormulaDecoder.Decode("=\"\""));
    }

    [Fact]
    public void Decode_PreservesEmbeddedDelimiters()
    {
        // The action separator (Chr(29)) and pipes ride inside the literal.
        var value = "a|script=s.py|input=A1\u001Db|script=t.py|input=B1";
        Assert.Equal(value, LegacyFormulaDecoder.Decode("=\"" + value + "\""));
    }

    [Theory]
    [InlineData("=A1:C10")]          // a real range reference, not a string literal
    [InlineData("=SUM(A1:A2)")]      // a formula
    [InlineData("=\"unterminated")]  // missing closing quote
    [InlineData("=\"a\"\"b")]        // trailing escaped quote leaves it open
    [InlineData("=\"a\"+\"b\"")]     // joined by something other than '&'
    [InlineData("=\"a\"\"")]         // dangling escaped quote, never closed
    public void Decode_NonStringLiteralFormula_ReturnsNull(string input)
    {
        Assert.Null(LegacyFormulaDecoder.Decode(input));
    }
}
