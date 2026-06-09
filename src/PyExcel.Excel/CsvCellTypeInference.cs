using System;
using System.Globalization;

namespace PyExcel.Excel;

/// <summary>
/// Per-cell type inference for CSV import — the logic that decides
/// whether a raw field string becomes a numeric / boolean / string cell.
/// Pure logic, cross-platform, used by the net48-only
/// <see cref="ImportService"/>; lives in its own class so the rules
/// (leading-zero guard, leading-plus guard, invariant-culture
/// double-parse) are unit-testable without Excel.
///
/// <para>The behaviour matches Excel's built-in CSV import: a string
/// that parses as a <see cref="double"/> under
/// <see cref="CultureInfo.InvariantCulture"/> becomes a numeric cell;
/// <c>"TRUE"</c> / <c>"FALSE"</c> (case-insensitive) become booleans;
/// everything else stays a string. Leading-zero strings (<c>"00123"</c>)
/// and leading-plus strings (<c>"+1234"</c>) intentionally fall through
/// to string-typed cells because parsing destroys formatting the user
/// likely cared about.</para>
/// </summary>
public static class CsvCellTypeInference
{
    /// <summary>Infer the cell value for one CSV field. Returns
    /// <see langword="null"/> for an empty field, <see cref="bool"/> for
    /// a recognised boolean token, <see cref="double"/> for a
    /// parseable number, and the original string otherwise.</summary>
    public static object? Infer(string? field)
    {
        if (string.IsNullOrEmpty(field)) return null;
        var s = field!;

        // Bool tokens, case-insensitive — matches what
        // ExportService.FormatCell emits for round-trip.
        if (string.Equals(s, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
        if (string.Equals(s, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;

        // Leading-zero guard: "0" alone is fine, but "00", "0123" preserve
        // formatting that double.Parse would lose. "0.5" should still
        // parse — the second char is '.', so we let it through.
        if (s.Length > 1 && s[0] == '0' && s[1] != '.') return s;

        // Leading-plus guard: "+1" might be a phone-number prefix etc.
        // Excel's CSV import keeps these as text.
        if (s[0] == '+') return s;

        if (double.TryParse(
                s,
                NumberStyles.Float | NumberStyles.AllowThousands,
                CultureInfo.InvariantCulture,
                out var d))
            return d;

        return s;
    }
}
