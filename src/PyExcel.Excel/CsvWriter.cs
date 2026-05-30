using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PyExcel.Excel;

/// <summary>
/// RFC-4180 CSV / TSV writer — Phase 5's export foundation. Produces a
/// text representation of a sequence of records that
/// <see cref="CsvParser"/> round-trips verbatim.
///
/// <para>Quoting rule (minimal, per RFC 4180): a field is wrapped in
/// double quotes iff it contains the delimiter, a double quote, a
/// carriage return, or a line feed. Internal double quotes are escaped
/// by doubling. Fields with leading or trailing whitespace are <em>not</em>
/// quoted unless they also contain one of the above — RFC 4180 §2.4
/// says spaces are part of the field and significant; quoting them
/// signals nothing extra.</para>
///
/// <para>Null cells render as the empty string. Other types are
/// rendered via <see cref="object.ToString"/>; the caller is responsible
/// for using a culture-invariant string representation for numerics and
/// dates if cross-locale stability matters.</para>
///
/// <para>The default line terminator is <c>CRLF</c> (RFC 4180 §2.1).
/// Callers that need <c>LF</c>-only output for Unix tooling can pass
/// <c>"\n"</c>.</para>
/// </summary>
public static class CsvWriter
{
    /// <summary>Convenience: write a tab-separated text representation
    /// (<c>.tsv</c> / Excel "Tab-separated text"). Identical to
    /// <see cref="Write(IEnumerable{IEnumerable{string?}}, char, string, bool)"/>
    /// with delimiter <c>'\t'</c>.</summary>
    public static string WriteTsv(IEnumerable<IEnumerable<string?>> rows)
        => Write(rows, '\t');

    /// <summary>Serialise <paramref name="rows"/> to a CSV string.</summary>
    /// <param name="rows">Records to write. Each record is a sequence of
    /// string fields; <see langword="null"/> fields render as the empty
    /// string. The sequence is enumerated once.</param>
    /// <param name="delimiter">Field separator. Defaults to comma.</param>
    /// <param name="lineTerminator">String to emit between records.
    /// Defaults to <c>CRLF</c> (RFC 4180); pass <c>"\n"</c> for Unix
    /// line endings.</param>
    /// <param name="writeBom">If <see langword="true"/>, prepend a UTF-8
    /// byte-order-mark character. Useful when the consumer is
    /// Excel-on-Windows opening the file by double-click, which uses the
    /// system codepage by default; the BOM forces UTF-8 interpretation.
    /// Defaults to <see langword="false"/>.</param>
    public static string Write(
        IEnumerable<IEnumerable<string?>> rows,
        char delimiter = ',',
        string lineTerminator = "\r\n",
        bool writeBom = false)
    {
        if (rows is null) throw new ArgumentNullException(nameof(rows));
        if (lineTerminator is null) throw new ArgumentNullException(nameof(lineTerminator));
        ValidateDelimiter(delimiter);

        var sb = new StringBuilder();
        if (writeBom) sb.Append('\uFEFF');

        var firstRow = true;
        foreach (var row in rows)
        {
            if (!firstRow) sb.Append(lineTerminator);
            firstRow = false;

            if (row is null)
            {
                // Treat a null record as an empty record (no fields).
                continue;
            }

            var firstField = true;
            foreach (var field in row)
            {
                if (!firstField) sb.Append(delimiter);
                firstField = false;
                AppendField(sb, field, delimiter);
            }
        }

        return sb.ToString();
    }

    /// <summary>Stream overload. Writes UTF-8-encoded CSV to
    /// <paramref name="stream"/>; the writer is flushed on exit but the
    /// stream is not closed.</summary>
    public static void Write(
        Stream stream,
        IEnumerable<IEnumerable<string?>> rows,
        char delimiter = ',',
        string lineTerminator = "\r\n",
        Encoding? encoding = null,
        bool writeBom = false)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (rows is null) throw new ArgumentNullException(nameof(rows));
        if (lineTerminator is null) throw new ArgumentNullException(nameof(lineTerminator));
        ValidateDelimiter(delimiter);

        // Default to UTF-8 without an emitted-by-the-encoder BOM; we
        // honour writeBom ourselves below so the caller has explicit
        // control regardless of which encoding object is passed.
        encoding ??= new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: false);

        using var writer = new StreamWriter(stream, encoding, bufferSize: 4096, leaveOpen: true);
        writer.NewLine = lineTerminator;

        if (writeBom) writer.Write('\uFEFF');

        var firstRow = true;
        foreach (var row in rows)
        {
            if (!firstRow) writer.Write(lineTerminator);
            firstRow = false;

            if (row is null) continue;

            var firstField = true;
            foreach (var field in row)
            {
                if (!firstField) writer.Write(delimiter);
                firstField = false;
                AppendField(writer, field, delimiter);
            }
        }

        writer.Flush();
    }

    private static void AppendField(StringBuilder sb, string? field, char delimiter)
    {
        if (string.IsNullOrEmpty(field)) return;
        if (NeedsQuoting(field!, delimiter))
        {
            sb.Append('"');
            foreach (var c in field!)
            {
                if (c == '"') sb.Append('"');
                sb.Append(c);
            }
            sb.Append('"');
        }
        else
        {
            sb.Append(field);
        }
    }

    private static void AppendField(StreamWriter writer, string? field, char delimiter)
    {
        if (string.IsNullOrEmpty(field)) return;
        if (NeedsQuoting(field!, delimiter))
        {
            writer.Write('"');
            foreach (var c in field!)
            {
                if (c == '"') writer.Write('"');
                writer.Write(c);
            }
            writer.Write('"');
        }
        else
        {
            writer.Write(field);
        }
    }

    private static bool NeedsQuoting(string field, char delimiter)
    {
        for (var i = 0; i < field.Length; i++)
        {
            var c = field[i];
            if (c == delimiter || c == '"' || c == '\r' || c == '\n')
                return true;
        }
        return false;
    }

    private static void ValidateDelimiter(char delimiter)
    {
        if (delimiter == '"' || delimiter == '\r' || delimiter == '\n')
            throw new ArgumentException(
                "Delimiter must not be a quote or newline character.",
                nameof(delimiter));
    }
}
