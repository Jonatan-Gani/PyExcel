using System;
using System.IO;

namespace PyExcel.Excel;

/// <summary>
/// Pure-logic planner for the Phase 5 Export button — validates the
/// user-typed ribbon fields and resolves the destination-file path
/// against the active workbook directory. Sibling of
/// <see cref="ImportPlanner"/>, kept separate so the symmetric
/// validation messages can phrase themselves around "source / target"
/// the way the user is reading the ribbon.
/// </summary>
public static class ExportPlanner
{
    /// <summary>
    /// Build the export plan from raw ribbon-field text. The
    /// <paramref name="workbookDirectory"/> is the active workbook's
    /// directory if it's saved; pass <see langword="null"/> for an
    /// unsaved workbook.
    /// </summary>
    /// <exception cref="FormatException">A required field is blank or
    /// the destination extension names a format we don't support yet
    /// (xlsx, xls, xlsm, xlsb, ods).</exception>
    public static ExportPlan Create(
        string? exportInput,
        string? exportOutput,
        string? workbookDirectory)
    {
        var source = (exportInput ?? string.Empty).Trim();
        if (source.Length == 0)
            throw new FormatException(
                "Export: the Input field is empty. " +
                "Type the source range (e.g. A1:C10, or Sheet1!A1:C10).");

        var target = (exportOutput ?? string.Empty).Trim();
        if (target.Length == 0)
            throw new FormatException(
                "Export: the Output field is empty. " +
                "Type the destination file path (.csv or .tsv).");

        var absolutePath = ImportPlanner.ResolvePath(target, workbookDirectory);
        var delimiter = ImportPlanner.DetectDelimiter(absolutePath);

        return new ExportPlan(source, absolutePath, delimiter);
    }
}

/// <summary>Validated, fully-resolved instructions for one export
/// operation.</summary>
public sealed record ExportPlan(
    string SourceRangeAddress,
    string AbsoluteTargetPath,
    char Delimiter);
