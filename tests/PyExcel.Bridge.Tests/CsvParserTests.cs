using System;
using System.IO;
using System.Linq;
using System.Text;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class CsvParserTests
{
    // -------------------------------------------------------------------------
    // Empty / trivial input
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_EmptyString_ReturnsEmpty()
    {
        Assert.Empty(CsvParser.Parse(string.Empty));
    }

    [Fact]
    public void Parse_NullText_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => CsvParser.Parse((string)null!));
    }

    [Fact]
    public void Parse_SingleField_OneRow()
    {
        var rows = CsvParser.Parse("a");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a" }, row);
    }

    [Fact]
    public void Parse_SingleRowThreeFields()
    {
        var rows = CsvParser.Parse("a,b,c");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b", "c" }, row);
    }

    // -------------------------------------------------------------------------
    // Line endings — RFC 4180 says CRLF, but real files mix all three.
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("a\r\nb")]
    [InlineData("a\nb")]
    [InlineData("a\rb")]
    public void Parse_LineEndings_AllVariants_TwoRows(string input)
    {
        var rows = CsvParser.Parse(input);
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "a" }, rows[0]);
        Assert.Equal(new[] { "b" }, rows[1]);
    }

    [Fact]
    public void Parse_MixedLineEndings_HandledTransparently()
    {
        var rows = CsvParser.Parse("a\r\nb\nc\rd");
        Assert.Equal(4, rows.Count);
        Assert.Equal(new[] { "a" }, rows[0]);
        Assert.Equal(new[] { "b" }, rows[1]);
        Assert.Equal(new[] { "c" }, rows[2]);
        Assert.Equal(new[] { "d" }, rows[3]);
    }

    [Fact]
    public void Parse_TrailingNewline_NoExtraEmptyRecord()
    {
        var rows = CsvParser.Parse("a,b\r\n");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b" }, row);
    }

    [Fact]
    public void Parse_TrailingLf_NoExtraEmptyRecord()
    {
        var rows = CsvParser.Parse("a,b\n");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b" }, row);
    }

    // -------------------------------------------------------------------------
    // Blank lines — Python's csv.reader emits empty records [] for these.
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_BlankLineBetweenRecords_EmitsEmptyRecord()
    {
        var rows = CsvParser.Parse("a\r\n\r\nb");
        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "a" }, rows[0]);
        Assert.Empty(rows[1]);
        Assert.Equal(new[] { "b" }, rows[2]);
    }

    [Fact]
    public void Parse_OnlyNewline_OneEmptyRecord()
    {
        var rows = CsvParser.Parse("\n");
        var row = Assert.Single(rows);
        Assert.Empty(row);
    }

    // -------------------------------------------------------------------------
    // BOM stripping
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_LeadingBom_Stripped()
    {
        var rows = CsvParser.Parse("\uFEFFa,b,c");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b", "c" }, row);
    }

    [Fact]
    public void Parse_NonLeadingBom_PreservedInField()
    {
        // A BOM mid-field is just a character; only the leading-byte
        // instance is stripped.
        var rows = CsvParser.Parse("a,b\uFEFFc");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b\uFEFFc" }, row);
    }

    // -------------------------------------------------------------------------
    // Quoting — escaped quotes, embedded delimiters, embedded newlines
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_QuotedField_WithCommaInside()
    {
        var rows = CsvParser.Parse("\"a,b\",c");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a,b", "c" }, row);
    }

    [Fact]
    public void Parse_QuotedField_WithEscapedQuote()
    {
        var rows = CsvParser.Parse("\"He said \"\"hi\"\"\",ok");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "He said \"hi\"", "ok" }, row);
    }

    [Fact]
    public void Parse_QuotedField_WithEmbeddedNewline()
    {
        var rows = CsvParser.Parse("\"line1\nline2\",b");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "line1\nline2", "b" }, row);
    }

    [Fact]
    public void Parse_QuotedField_WithEmbeddedCrlf()
    {
        var rows = CsvParser.Parse("\"l1\r\nl2\",b");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "l1\r\nl2", "b" }, row);
    }

    [Fact]
    public void Parse_EmptyQuotedField()
    {
        var rows = CsvParser.Parse("\"\",a,\"\"");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "", "a", "" }, row);
    }

    [Fact]
    public void Parse_AllQuoted()
    {
        var rows = CsvParser.Parse("\"a\",\"b\",\"c\"");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b", "c" }, row);
    }

    // -------------------------------------------------------------------------
    // Empty fields — leading, trailing, consecutive delimiters
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_LeadingDelimiter_EmptyFirstField()
    {
        var rows = CsvParser.Parse(",a,b");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "", "a", "b" }, row);
    }

    [Fact]
    public void Parse_TrailingDelimiter_EmptyLastField()
    {
        var rows = CsvParser.Parse("a,b,");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b", "" }, row);
    }

    [Fact]
    public void Parse_ConsecutiveDelimiters_EmptyMiddleField()
    {
        var rows = CsvParser.Parse("a,,b");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "", "b" }, row);
    }

    [Fact]
    public void Parse_OnlyDelimiter_TwoEmptyFields()
    {
        var rows = CsvParser.Parse(",");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "", "" }, row);
    }

    // -------------------------------------------------------------------------
    // Whitespace — preserved per RFC 4180 §2.4
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_LeadingSpaces_Preserved()
    {
        var rows = CsvParser.Parse("  a,  b");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "  a", "  b" }, row);
    }

    [Fact]
    public void Parse_TrailingSpaces_Preserved()
    {
        var rows = CsvParser.Parse("a  ,b  ");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a  ", "b  " }, row);
    }

    // -------------------------------------------------------------------------
    // Permissive: bare quotes, content after closing quote
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_BareQuoteInUnquotedField_PassedThrough()
    {
        var rows = CsvParser.Parse("a\"b,c");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a\"b", "c" }, row);
    }

    [Fact]
    public void Parse_ContentAfterClosingQuote_Appended()
    {
        // Matches Python's csv.reader default mode and Excel's import.
        var rows = CsvParser.Parse("\"a\"b,c");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "ab", "c" }, row);
    }

    // -------------------------------------------------------------------------
    // Errors
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_UnterminatedQuote_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => CsvParser.Parse("\"hello"));
        Assert.Contains("Unterminated", ex.Message);
    }

    [Fact]
    public void Parse_UnterminatedQuoteAcrossLines_Throws()
    {
        Assert.Throws<FormatException>(() => CsvParser.Parse("\"hello\nworld"));
    }

    [Theory]
    [InlineData('"')]
    [InlineData('\r')]
    [InlineData('\n')]
    public void Parse_InvalidDelimiter_Throws(char delimiter)
    {
        Assert.Throws<ArgumentException>(() => CsvParser.Parse("a,b", delimiter));
    }

    // -------------------------------------------------------------------------
    // Alternative delimiters — TSV, semicolon (European Excel)
    // -------------------------------------------------------------------------

    [Fact]
    public void ParseTsv_TabDelimited()
    {
        var rows = CsvParser.ParseTsv("a\tb\tc\nd\te\tf");
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "a", "b", "c" }, rows[0]);
        Assert.Equal(new[] { "d", "e", "f" }, rows[1]);
    }

    [Fact]
    public void Parse_TabAsDelimiter_CommasArePartOfField()
    {
        // With tab as delimiter, commas inside fields are not special.
        var rows = CsvParser.Parse("1,000\t2,000\n3,000\t4,000", '\t');
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "1,000", "2,000" }, rows[0]);
        Assert.Equal(new[] { "3,000", "4,000" }, rows[1]);
    }

    [Fact]
    public void Parse_SemicolonDelimiter()
    {
        var rows = CsvParser.Parse("a;b;c", ';');
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b", "c" }, row);
    }

    [Fact]
    public void Parse_TsvWithQuotedField_HandlesEmbeddedTab()
    {
        var rows = CsvParser.ParseTsv("\"a\tb\"\tc");
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a\tb", "c" }, row);
    }

    // -------------------------------------------------------------------------
    // Stream overload
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_Stream_DefaultUtf8()
    {
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes("a,b\r\nc,d"));
        var rows = CsvParser.Parse(ms);
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "a", "b" }, rows[0]);
        Assert.Equal(new[] { "c", "d" }, rows[1]);
    }

    [Fact]
    public void Parse_Stream_WithUtf8BomBytes_BomStripped()
    {
        // UTF-8 BOM bytes: EF BB BF, then "a,b".
        var bytes = new byte[] { 0xEF, 0xBB, 0xBF }
            .Concat(Encoding.UTF8.GetBytes("a,b"))
            .ToArray();
        using var ms = new MemoryStream(bytes);
        var rows = CsvParser.Parse(ms);
        var row = Assert.Single(rows);
        Assert.Equal(new[] { "a", "b" }, row);
    }

    [Fact]
    public void Parse_Stream_NullStream_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => CsvParser.Parse((Stream)null!));
    }

    [Fact]
    public void Parse_Stream_LeavesStreamOpen()
    {
        using var ms = new MemoryStream(Encoding.UTF8.GetBytes("a,b"));
        CsvParser.Parse(ms);
        // If the stream had been closed by Parse, this would throw.
        Assert.True(ms.CanRead);
    }

    // -------------------------------------------------------------------------
    // Multi-row, multi-field — realistic Excel export
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_RealisticSample_Roundtrip()
    {
        const string input =
            "Name,Quantity,Notes\r\n" +
            "Widget,42,\"Default size\"\r\n" +
            "Gizmo,7,\"Has \"\"special\"\" handling\"\r\n" +
            "Doodad,0,\r\n";
        var rows = CsvParser.Parse(input);
        Assert.Equal(4, rows.Count);
        Assert.Equal(new[] { "Name", "Quantity", "Notes" }, rows[0]);
        Assert.Equal(new[] { "Widget", "42", "Default size" }, rows[1]);
        Assert.Equal(new[] { "Gizmo", "7", "Has \"special\" handling" }, rows[2]);
        Assert.Equal(new[] { "Doodad", "0", "" }, rows[3]);
    }
}
