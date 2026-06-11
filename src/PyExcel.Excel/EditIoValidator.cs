using System;

namespace PyExcel.Excel;

/// <summary>
/// The outcome of validating an Edit-Import / Edit-Export / Edit-Paste
/// dialog: either the trimmed field values to persist back to the
/// workbook state, or an inline error for the dialog to show without
/// closing.
/// </summary>
public sealed class EditIoValidationResult
{
    private EditIoValidationResult(bool isValid, string? error, string? input, string? output)
    {
        IsValid = isValid;
        ErrorMessage = error;
        Input = input;
        Output = output;
    }

    public bool IsValid { get; }
    public string? ErrorMessage { get; }

    /// <summary>The trimmed input field to persist (the import source path
    /// / export source range). Null for Paste, which has no input.</summary>
    public string? Input { get; }

    /// <summary>The trimmed output field to persist (the import target
    /// range / export target path / paste target range).</summary>
    public string? Output { get; }

    internal static EditIoValidationResult Ok(string? input, string? output)
        => new(true, null, input, output);

    internal static EditIoValidationResult Fail(string error)
        => new(false, error, null, null);
}

/// <summary>
/// Pure validation for the Edit-Import / Edit-Export / Edit-Paste dialogs
/// (Phase 8). Each reuses the same planner the run-time service uses
/// (<see cref="ImportPlanner"/> / <see cref="ExportPlanner"/>) as the
/// validity check, so the dialog rejects exactly what the service would —
/// blank fields, unsupported formats (.xls / .ods / Excel export target),
/// invalid paths — with the same messages, before the value is saved.
/// Kept cross-platform so it's unit-tested on Linux CI.
/// </summary>
public static class EditIoValidator
{
    /// <summary>Validate the Edit-Import fields (source file → target
    /// range). Returns the trimmed values to persist, or an inline
    /// message.</summary>
    public static EditIoValidationResult ValidateImport(
        string? input, string? output, string? workbookDirectory)
    {
        try
        {
            ImportPlanner.Create(input, output, workbookDirectory);
        }
        catch (FormatException ex) { return EditIoValidationResult.Fail(ex.Message); }
        catch (ArgumentException ex)
        {
            return EditIoValidationResult.Fail($"Import: {ex.Message}");
        }

        return EditIoValidationResult.Ok(
            (input ?? string.Empty).Trim(), (output ?? string.Empty).Trim());
    }

    /// <summary>Validate the Edit-Export fields (source range → target
    /// file). Returns the trimmed values to persist, or an inline
    /// message.</summary>
    public static EditIoValidationResult ValidateExport(
        string? input, string? output, string? workbookDirectory)
    {
        try
        {
            ExportPlanner.Create(input, output, workbookDirectory);
        }
        catch (FormatException ex) { return EditIoValidationResult.Fail(ex.Message); }
        catch (ArgumentException ex)
        {
            return EditIoValidationResult.Fail($"Export: {ex.Message}");
        }

        return EditIoValidationResult.Ok(
            (input ?? string.Empty).Trim(), (output ?? string.Empty).Trim());
    }

    /// <summary>Validate the Edit-Paste field (target range only — Paste
    /// pulls its data from the last archived run, not a user field).
    /// Returns the trimmed target range to persist, or an inline
    /// message.</summary>
    public static EditIoValidationResult ValidatePaste(string? output)
    {
        var target = (output ?? string.Empty).Trim();
        if (target.Length == 0)
            return EditIoValidationResult.Fail(
                "Paste: the target range is empty. " +
                "Type the range to paste into (e.g. A1, or Sheet1!A1).");

        return EditIoValidationResult.Ok(null, target);
    }
}
