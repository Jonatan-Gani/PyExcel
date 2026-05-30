using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class CsvWriterTests
{
    // -------------------------------------------------------------------------
    // Trivial cases
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_NullRows_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => CsvWriter.Write(null!));
    }

    [Fact]
    public void Write_NullLineTerminator_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => CsvWriter.Write(new[] { new[] { "a" } }, lineTerminator: null!));
    }

    [Fact]
    public void Write_EmptyRows_ReturnsEmpty()
    {
        var s = CsvWriter.Write(Array.Empty<IEnumerable<string?>>());
        Assert.Equal(string.Empty, s);
    }

    [Fact]
    public void Write_SingleRowSingleField()
    {
        var s = CsvWriter.Write(new[] { new[] { "a" } });
        Assert.Equal("a", s);
    }

    // -------------------------------------------------------------------------
    // Quoting — minimal per RFC 4180
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_PlainFields_NoQuoting()
    {
        var s = CsvWriter.Write(new[] { new[] { "a", "b", "c" } });
        Assert.Equal("a,b,c", s);
    }

    [Fact]
    public void Write_FieldWithComma_Quoted()
    {
        var s = CsvWriter.Write(new[] { new[] { "a,b", "c" } });
        Assert.Equal("\"a,b\",c", s);
    }

    [Fact]
    public void Write_FieldWithQuote_QuotedAndEscaped()
    {
        var s = CsvWriter.Write(new[] { new[] { "He said \"hi\"" } });
        Assert.Equal("\"He said \"\"hi\"\"\"", s);
    }

    [Fact]
    public void Write_FieldWithLf_Quoted()
    {
        var s = CsvWriter.Write(new[] { new[] { "line1\nline2", "b" } });
        Assert.Equal("\"line1\nline2\",b", s);
    }

    [Fact]
    public void Write_FieldWithCr_Quoted()
    {
        var s = CsvWriter.Write(new[] { new[] { "line1\rline2", "b" } });
        Assert.Equal("\"line1\rline2\",b", s);
    }

    [Fact]
    public void Write_FieldWithSpaces_NotQuoted()
    {
        // RFC 4180 §2.4 — spaces are significant but don't trigger quoting.
        var s = CsvWriter.Write(new[] { new[] { "  a  ", " b " } });
        Assert.Equal("  a  , b ", s);
    }

    [Fact]
    public void Write_EmptyField_NoQuoting()
    {
        var s = CsvWriter.Write(new[] { new[] { "", "a", "" } });
        Assert.Equal(",a,", s);
    }

    [Fact]
    public void Write_NullField_RendersEmpty()
    {
        var s = CsvWriter.Write(new[] { new string?[] { null, "a", null } });
        Assert.Equal(",a,", s);
    }

    // -------------------------------------------------------------------------
    // Line terminators
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_MultipleRows_CrlfByDefault()
    {
        var s = CsvWriter.Write(new[]
        {
            new[] { "a", "b" },
            new[] { "c", "d" },
        });
        Assert.Equal("a,b\r\nc,d", s);
    }

    [Fact]
    public void Write_MultipleRows_LfOverride()
    {
        var s = CsvWriter.Write(new[]
        {
            new[] { "a", "b" },
            new[] { "c", "d" },
        }, lineTerminator: "\n");
        Assert.Equal("a,b\nc,d", s);
    }

    [Fact]
    public void Write_NoTrailingNewline()
    {
        // RFC 4180: optional. We do not emit one.
        var s = CsvWriter.Write(new[] { new[] { "a" } });
        Assert.Equal("a", s);
    }

    [Fact]
    public void Write_TwoRowsLastEmpty_TrailingSeparatorOnly()
    {
        // [["a"], []] → "a\r\n"
        var s = CsvWriter.Write(new[]
        {
            new[] { "a" },
            Array.Empty<string>(),
        });
        Assert.Equal("a\r\n", s);
    }

    // -------------------------------------------------------------------------
    // BOM option
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_WithBom_PrependsBomChar()
    {
        var s = CsvWriter.Write(new[] { new[] { "a", "b" } }, writeBom: true);
        Assert.Equal('\uFEFF', s[0]);
        Assert.Equal("\uFEFFa,b", s);
    }

    [Fact]
    public void Write_DefaultNoBom()
    {
        var s = CsvWriter.Write(new[] { new[] { "a", "b" } });
        Assert.NotEqual('\uFEFF', s[0]);
    }

    // -------------------------------------------------------------------------
    // Tab / semicolon delimiters
    // -------------------------------------------------------------------------

    [Fact]
    public void WriteTsv_TabDelimited()
    {
        var s = CsvWriter.WriteTsv(new[] { new[] { "a", "b", "c" } });
        Assert.Equal("a\tb\tc", s);
    }

    [Fact]
    public void Write_TabDelimited_CommaNotQuoted()
    {
        // Comma is no longer the delimiter, so a field containing a comma
        // doesn't need quoting.
        var s = CsvWriter.Write(new[] { new[] { "1,000", "2,000" } }, '\t');
        Assert.Equal("1,000\t2,000", s);
    }

    [Fact]
    public void Write_TabDelimited_TabFieldQuoted()
    {
        var s = CsvWriter.Write(new[] { new[] { "a\tb", "c" } }, '\t');
        Assert.Equal("\"a\tb\"\tc", s);
    }

    [Fact]
    public void Write_SemicolonDelimited()
    {
        var s = CsvWriter.Write(new[] { new[] { "a", "b" } }, ';');
        Assert.Equal("a;b", s);
    }

    [Theory]
    [InlineData('"')]
    [InlineData('\r')]
    [InlineData('\n')]
    public void Write_InvalidDelimiter_Throws(char delimiter)
    {
        Assert.Throws<ArgumentException>(
            () => CsvWriter.Write(new[] { new[] { "a" } }, delimiter));
    }

    // -------------------------------------------------------------------------
    // Stream overload
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_Stream_DefaultUtf8()
    {
        using var ms = new MemoryStream();
        CsvWriter.Write(ms, new[] { new[] { "a", "b" }, new[] { "c", "d" } });
        var text = Encoding.UTF8.GetString(ms.ToArray());
        Assert.Equal("a,b\r\nc,d", text);
    }

    [Fact]
    public void Write_Stream_WithBom_EmitsUtf8BomBytes()
    {
        using var ms = new MemoryStream();
        CsvWriter.Write(ms, new[] { new[] { "a" } }, writeBom: true);
        var bytes = ms.ToArray();
        // UTF-8 BOM is EF BB BF.
        Assert.Equal(0xEF, bytes[0]);
        Assert.Equal(0xBB, bytes[1]);
        Assert.Equal(0xBF, bytes[2]);
    }

    [Fact]
    public void Write_Stream_LeavesStreamOpen()
    {
        using var ms = new MemoryStream();
        CsvWriter.Write(ms, new[] { new[] { "a" } });
        // If the stream had been closed by Write, this would throw.
        Assert.True(ms.CanWrite);
    }

    // -------------------------------------------------------------------------
    // Round-trip — parser ↔ writer on tricky payloads
    // -------------------------------------------------------------------------

    [Fact]
    public void Roundtrip_QuotedField_WithCommaQuoteNewline()
    {
        var original = new[]
        {
            new[] { "a,b", "He said \"hi\"", "line1\nline2" },
            new[] { "plain", "", " trailing " },
        };
        var encoded = CsvWriter.Write(original);
        var decoded = CsvParser.Parse(encoded);
        Assert.Equal(original.Length, decoded.Count);
        for (var r = 0; r < original.Length; r++)
        {
            Assert.Equal(original[r], decoded[r]);
        }
    }

    [Fact]
    public void Roundtrip_Tsv()
    {
        var original = new[]
        {
            new[] { "a", "b\tinside", "c" },
            new[] { "d", "e", "f,with,commas" },
        };
        var encoded = CsvWriter.WriteTsv(original);
        var decoded = CsvParser.ParseTsv(encoded);
        Assert.Equal(original.Length, decoded.Count);
        for (var r = 0; r < original.Length; r++)
        {
            Assert.Equal(original[r], decoded[r]);
        }
    }
}
