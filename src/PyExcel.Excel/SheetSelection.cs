using System;
using System.Collections.Generic;

namespace PyExcel.Excel;

/// <summary>How an Excel-import sheet selection resolves.</summary>
public enum SheetResolutionKind
{
    /// <summary>A specific sheet is determined — either the user pinned one
    /// with the <c>path!Sheet</c> syntax, or the workbook has exactly one
    /// sheet so there's nothing to choose.</summary>
    Resolved,

    /// <summary>No sheet was pinned and the workbook has more than one, so
    /// the user should be asked which to import.</summary>
    Prompt,

    /// <summary>The workbook has no sheets to import.</summary>
    Empty,
}

/// <summary>
/// The decision of which sheet an Excel import should read, given the
/// user's (optional) pinned sheet and the workbook's actual sheets.
/// </summary>
public readonly struct SheetResolution
{
    private SheetResolution(SheetResolutionKind kind, string? sheet, IReadOnlyList<string> available)
    {
        Kind = kind;
        Sheet = sheet;
        AvailableSheets = available;
    }

    public SheetResolutionKind Kind { get; }

    /// <summary>The resolved sheet name when <see cref="Kind"/> is
    /// <see cref="SheetResolutionKind.Resolved"/>; otherwise null.</summary>
    public string? Sheet { get; }

    /// <summary>The sheets to offer when <see cref="Kind"/> is
    /// <see cref="SheetResolutionKind.Prompt"/>; otherwise empty.</summary>
    public IReadOnlyList<string> AvailableSheets { get; }

    internal static SheetResolution Resolved(string sheet)
        => new(SheetResolutionKind.Resolved, sheet, Array.Empty<string>());

    internal static SheetResolution Prompt(IReadOnlyList<string> available)
        => new(SheetResolutionKind.Prompt, null, available);

    internal static SheetResolution NoSheets()
        => new(SheetResolutionKind.Empty, null, Array.Empty<string>());
}

/// <summary>
/// Pure decision logic for which sheet an Excel import reads — kept apart
/// from the COM-bound <c>ImportService</c> so it's unit-testable on Linux
/// CI. The service consults this after opening the workbook and
/// enumerating its sheets, then either reads the resolved sheet or shows
/// the sheet picker for the prompt case.
/// </summary>
public static class SheetSelection
{
    /// <summary>
    /// Decide the sheet from the user's pinned sheet (the <c>!Sheet</c>
    /// suffix, may be null) and the workbook's sheet names.
    /// <list type="bullet">
    /// <item>A non-blank pinned sheet always wins — <see cref="SheetResolutionKind.Resolved"/>
    /// (a missing name is surfaced later by the COM lookup, not here).</item>
    /// <item>No pin + one sheet → that sheet, <see cref="SheetResolutionKind.Resolved"/>.</item>
    /// <item>No pin + several sheets → <see cref="SheetResolutionKind.Prompt"/>.</item>
    /// <item>No sheets at all → <see cref="SheetResolutionKind.Empty"/>.</item>
    /// </list>
    /// </summary>
    public static SheetResolution Resolve(string? specifiedSheet, IReadOnlyList<string> availableSheets)
    {
        if (availableSheets is null)
            throw new ArgumentNullException(nameof(availableSheets));

        var pinned = (specifiedSheet ?? string.Empty).Trim();
        if (pinned.Length > 0)
            return SheetResolution.Resolved(pinned);

        if (availableSheets.Count == 0)
            return SheetResolution.NoSheets();
        if (availableSheets.Count == 1)
            return SheetResolution.Resolved(availableSheets[0]);

        return SheetResolution.Prompt(availableSheets);
    }
}
