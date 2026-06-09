using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PyExcel.Excel;

/// <summary>
/// RFC-4180 CSV / TSV parser — Phase 5's import foundation. Produces a list
/// of records, each a list of field strings, in source order.
///
/// <para>Conformance: the parser implements RFC 4180 plus the common Excel
/// extensions that real-world CSV depends on:</para>
/// <list type="bullet">
///   <item>Records may be terminated by <c>CRLF</c>, <c>LF</c>, or bare
///     <c>CR</c>. Mixed terminators within one file are tolerated.</item>
///   <item>A UTF-8 BOM at the start of the input is consumed silently.</item>
///   <item>Fields enclosed in double quotes may contain the delimiter,
///     newlines, and escaped quotes (<c>""</c>). Quoting outside of fields
///     enclosed at field start is permissive — content after a closing
///     quote and before the next delimiter is appended to the field,
///     matching Python's <c>csv.reader</c> and Excel's import behaviour.</item>
///   <item>A blank line between records becomes an empty record (<c>[]</c>),
///     matching Python's <c>csv.reader</c>. A trailing blank line at EOF is
///     not emitted as an empty record — it's recognised as the terminator
///     of the previous record.</item>
///   <item>Spaces inside a field are preserved verbatim. RFC 4180 §2.4
///     says spaces are part of the field; we do not trim.</item>
/// </list>
///
/// <para>Pure logic — no IO, no globals, no locale dependency. The
/// <see cref="Parse(Stream, char, Encoding?)"/> overload reads a stream,
/// stripping a UTF-8 BOM if present and decoding under the supplied
/// encoding (default <see cref="Encoding.UTF8"/>).</para>
///
/// <para>Errors: an unterminated quoted field (EOF inside a quote) throws
/// <see cref="FormatException"/>. A delimiter equal to <c>'"'</c>,
/// <c>'\r'</c>, or <c>'\n'</c> is rejected via
/// <see cref="ArgumentException"/> — those characters are structural.</para>
/// </summary>
public static class CsvParser
{
    /// <summary>UTF-8 byte-order-mark character (U+FEFF). We strip a
    /// single leading instance silently — many Excel exports prepend
    /// one.</summary>
    private const char Bom = '\uFEFF';

    /// <summary>Convenience: parse a tab-separated file (<c>.tsv</c> /
    /// Excel "Tab-separated text"). Identical to
    /// <see cref="Parse(string, char)"/> with delimiter <c>'\t'</c>.</summary>
    public static IReadOnlyList<IReadOnlyList<string>> ParseTsv(string text)
        => Parse(text, '\t');

    /// <summary>Parse CSV / TSV text into a list of records.</summary>
    /// <param name="text">The full input text. Must not be <see langword="null"/>.</param>
    /// <param name="delimiter">Field separator. Defaults to comma; pass
    /// <c>'\t'</c> for TSV or <c>';'</c> for European-locale Excel exports.</param>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="delimiter"/> is
    /// a quote or newline character (structurally invalid).</exception>
    /// <exception cref="FormatException">An opening quote is never closed
    /// before end of input.</exception>
    public static IReadOnlyList<IReadOnlyList<string>> Parse(string text, char delimiter = ',')
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        ValidateDelimiter(delimiter);

        var rows = new List<IReadOnlyList<string>>();
        var record = new List<string>();
        var field = new StringBuilder();
        var inQuotes = false;
        // True iff the current field has any captured content yet —
        // either characters appended to the buffer, or a recognised
        // opening quote. Used to distinguish a field-start quote
        // (opener) from a literal quote inside an unquoted field.
        var fieldStarted = false;
        // True iff the current record has any committed field OR any
        // in-progress field content. A line break with no row content
        // emits an empty record (matches Python's csv.reader).
        var rowHasContent = false;

        var i = 0;
        // Strip a single leading UTF-8 BOM if present.
        if (text.Length > 0 && text[0] == Bom) i = 1;

        while (i < text.Length)
        {
            var c = text[i];

            if (inQuotes)
            {
                if (c == '"')
                {
                    // "" inside a quoted field = literal quote.
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        field.Append('"');
                        i += 2;
                    }
                    else
                    {
                        // End of the quoted run. Further characters until
                        // the next delimiter or newline are appended to
                        // the same field (Excel/Python permissive mode).
                        inQuotes = false;
                        i++;
                    }
                }
                else
                {
                    field.Append(c);
                    i++;
                }
                continue;
            }

            // Not in quotes.
            if (c == '"')
            {
                if (!fieldStarted)
                {
                    // Opening quote — enter quoted mode.
                    inQuotes = true;
                    fieldStarted = true;
                    rowHasContent = true;
                    i++;
                }
                else
                {
                    // Bare quote in the middle of an unquoted field, or
                    // after a closed quoted run. Permissive: pass through.
                    field.Append(c);
                    i++;
                }
            }
            else if (c == delimiter)
            {
                record.Add(field.ToString());
                field.Clear();
                fieldStarted = false;
                rowHasContent = true;
                i++;
            }
            else if (c == '\r' || c == '\n')
            {
                // Consume CRLF as a single terminator.
                if (c == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i += 2;
                else
                    i++;

                if (rowHasContent || fieldStarted)
                {
                    record.Add(field.ToString());
                    field.Clear();
                    rows.Add(record);
                    record = new List<string>();
                }
                else
                {
                    // Blank line — emit empty record, do not commit a
                    // single empty field.
                    rows.Add(record);
                    record = new List<string>();
                }
                fieldStarted = false;
                rowHasContent = false;
            }
            else
            {
                field.Append(c);
                fieldStarted = true;
                rowHasContent = true;
                i++;
            }
        }

        if (inQuotes)
            throw new FormatException(
                "Unterminated quoted field at end of input.");

        // Final record — only emit if any field content was seen. A
        // trailing line terminator at EOF leaves rowHasContent=false and
        // produces no extra empty record.
        if (rowHasContent || fieldStarted)
        {
            record.Add(field.ToString());
            rows.Add(record);
        }

        return rows;
    }

    /// <summary>Parse CSV / TSV from a byte stream. Decodes under
    /// <paramref name="encoding"/> (default <see cref="Encoding.UTF8"/>);
    /// the leading UTF-8 BOM, if present, is stripped after decoding.</summary>
    /// <param name="stream">Stream positioned at the start of the CSV
    /// payload. Read to end; not closed.</param>
    public static IReadOnlyList<IReadOnlyList<string>> Parse(
        Stream stream, char delimiter = ',', Encoding? encoding = null)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        ValidateDelimiter(delimiter);
        encoding ??= new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);

        // detectEncodingFromByteOrderMarks=true lets the reader pick up a
        // UTF-16 / UTF-8 BOM and use the matching codec. Even when the
        // caller passes a specific encoding we still let the BOM win —
        // that's what Excel does on import, and it's the safer default.
        using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);
        var text = reader.ReadToEnd();
        return Parse(text, delimiter);
    }

    private static void ValidateDelimiter(char delimiter)
    {
        if (delimiter == '"' || delimiter == '\r' || delimiter == '\n')
            throw new ArgumentException(
                "Delimiter must not be a quote or newline character.",
                nameof(delimiter));
    }
}
