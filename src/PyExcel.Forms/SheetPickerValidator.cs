using System;
using System.Collections.Generic;

namespace PyExcel.Forms;

/// <summary>
/// The outcome of validating the sheet-picker dialog's selection: either
/// the chosen sheet name (canonical casing, as it appears in the
/// workbook) or an inline error for the dialog to show without closing.
/// </summary>
public sealed class SheetPickerValidationResult
{
    private SheetPickerValidationResult(bool isValid, string? error, string? sheet)
    {
        IsValid = isValid;
        ErrorMessage = error;
        SelectedSheet = sheet;
    }

    public bool IsValid { get; }
    public string? ErrorMessage { get; }

    /// <summary>The chosen sheet, in the casing the workbook reports — null
    /// when not <see cref="IsValid"/>.</summary>
    public string? SelectedSheet { get; }

    internal static SheetPickerValidationResult Ok(string sheet)
        => new(true, null, sheet);

    internal static SheetPickerValidationResult Fail(string error)
        => new(false, error, null);
}

/// <summary>
/// Pure validation for the sheet-picker dialog (the Phase 8 WinForms port
/// of v1's <c>SheetPickerForm</c>). Kept cross-platform so the dialog's
/// "a sheet must be chosen, and it must be one the workbook actually
/// offers" rule is unit-tested on Linux CI without Excel.
/// </summary>
public static class SheetPickerValidator
{
    /// <summary>
    /// Validate the picker's current selection against the offered sheets.
    /// Membership is case-insensitive (Excel sheet names are unique
    /// case-insensitively, and <c>Workbook.Sheets[name]</c> looks up that
    /// way); the canonical name from <paramref name="availableSheets"/> is
    /// returned so the caller composes the reference with the workbook's
    /// own casing.
    /// </summary>
    public static SheetPickerValidationResult Validate(
        string? selected,
        IEnumerable<string> availableSheets)
    {
        if (availableSheets is null)
            throw new ArgumentNullException(nameof(availableSheets));

        var trimmed = (selected ?? string.Empty).Trim();
        if (trimmed.Length == 0)
            return SheetPickerValidationResult.Fail("Select a sheet to import.");

        foreach (var sheet in availableSheets)
        {
            if (string.Equals(sheet, trimmed, StringComparison.OrdinalIgnoreCase))
                return SheetPickerValidationResult.Ok(sheet);
        }

        return SheetPickerValidationResult.Fail(
            $"Sheet '{trimmed}' is not in this workbook.");
    }
}
