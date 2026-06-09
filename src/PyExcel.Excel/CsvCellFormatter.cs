using System;
using System.Globalization;

namespace PyExcel.Excel;

/// <summary>
/// Cell-to-string formatting for CSV export — the symmetric counterpart
/// to <see cref="CsvCellTypeInference"/>. Pure logic, cross-platform;
/// the net48-only <see cref="ExportService"/> calls into it after
/// stripping Excel-DNA's <c>ExcelEmpty</c> / <c>ExcelMissing</c> /
/// <c>ExcelError</c> sentinels (which need an ExcelDna reference
/// available only on net48).
///
/// <para>Rules:</para>
/// <list type="bullet">
///   <item><see langword="null"/> → <see langword="null"/> (which
///     <see cref="CsvWriter"/> renders as the empty string).</item>
///   <item><see cref="double"/> → invariant-culture round-trip format
///     (<c>"R"</c>) so re-importing yields the same value bit-for-bit.</item>
///   <item><see cref="bool"/> → <c>"TRUE"</c> / <c>"FALSE"</c> matching
///     Excel's display convention and
///     <see cref="CsvCellTypeInference"/>'s recognition.</item>
///   <item><see cref="DateTime"/> → ISO 8601
///     (<c>"yyyy-MM-ddTHH:mm:ss"</c>) so the export is
///     locale-independent.</item>
///   <item><see cref="string"/> → pass-through.</item>
///   <item>Anything else → <see cref="Convert.ToString(object?, IFormatProvider?)"/>
///     under invariant culture.</item>
/// </list>
/// </summary>
public static class CsvCellFormatter
{
    /// <summary>Render an Excel-side cell value as the string CSV
    /// expects. Returns <see langword="null"/> for null input — the
    /// writer treats null as the empty field.</summary>
    public static string? Format(object? value)
    {
        if (value is null) return null;
        switch (value)
        {
            case double d:
                return d.ToString("R", CultureInfo.InvariantCulture);
            case bool b:
                return b ? "TRUE" : "FALSE";
            case DateTime dt:
                return dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
            case string s:
                return s;
            default:
                return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }
}
