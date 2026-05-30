using System;
using System.IO;

namespace PyExcel.Excel;

/// <summary>
/// Pure-logic planner for the Phase 5 Import button — validates the
/// user-typed ribbon fields, resolves the source-file path against the
/// active workbook directory, and detects the field delimiter from the
/// file extension.
///
/// <para>Separated from the COM-bound <c>ImportService</c> so its logic
/// (path resolution rules, delimiter detection, validation messages) is
/// unit-testable on Linux CI without needing Excel.</para>
///
/// <para>Errors are <see cref="FormatException"/> for malformed input
/// (missing field, unsupported extension) and
/// <see cref="ArgumentException"/> for structurally-invalid arguments.
/// The service layer catches these and surfaces them to LogDisplay.</para>
/// </summary>
public static class ImportPlanner
{
    /// <summary>
    /// Build the import plan from raw ribbon-field text. The
    /// <paramref name="workbookDirectory"/> is the active workbook's
    /// directory if it's saved; pass <see langword="null"/> for an
    /// unsaved workbook (relative paths then resolve against the process
    /// working directory).
    /// </summary>
    /// <exception cref="FormatException">A required field is blank or
    /// the file extension names a format we don't support yet (xlsx,
    /// xls, xlsm, xlsb, ods — those are the Excel-format-import
    /// follow-up item).</exception>
    public static ImportPlan Create(
        string? importInput,
        string? importOutput,
        string? workbookDirectory)
    {
        var source = (importInput ?? string.Empty).Trim();
        if (source.Length == 0)
            throw new FormatException(
                "Import: the Input field is empty. " +
                "Type the path to a CSV / TSV file to import from.");

        var target = (importOutput ?? string.Empty).Trim();
        if (target.Length == 0)
            throw new FormatException(
                "Import: the Output field is empty. " +
                "Type the range to write into (e.g. A1, or Sheet1!A1).");

        var absolutePath = ResolvePath(source, workbookDirectory);
        var delimiter = DetectDelimiter(absolutePath);

        return new ImportPlan(absolutePath, delimiter, target);
    }

    /// <summary>Resolve a (possibly relative) source path against the
    /// workbook directory. Absolute paths pass through; rooted-but-no-drive
    /// paths (like <c>/foo</c> on Windows) also pass through because
    /// <see cref="Path.IsPathRooted(string)"/> says so. Relative paths
    /// resolve via <see cref="Path.Combine(string,string)"/> with
    /// <see cref="Path.GetFullPath(string)"/> normalising the result.</summary>
    public static string ResolvePath(string source, string? workbookDirectory)
    {
        if (string.IsNullOrWhiteSpace(source))
            throw new ArgumentException("source path is blank", nameof(source));

        if (Path.IsPathRooted(source))
            return Path.GetFullPath(source);

        var basis = string.IsNullOrWhiteSpace(workbookDirectory)
            ? Environment.CurrentDirectory
            : workbookDirectory!;
        return Path.GetFullPath(Path.Combine(basis, source));
    }

    /// <summary>Detect the field delimiter from a file extension. The
    /// rule is intentionally narrow: <c>.tsv</c> → tab, everything else
    /// (including no extension) → comma. Binary spreadsheet formats
    /// (<c>.xlsx</c> / <c>.xls</c> / <c>.xlsm</c> / <c>.xlsb</c> /
    /// <c>.ods</c>) throw <see cref="FormatException"/> — those need
    /// the separate Excel-format-import service that's tracked as a
    /// Phase 5 follow-up.</summary>
    public static char DetectDelimiter(string path)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        var ext = Path.GetExtension(path).ToLowerInvariant();
        switch (ext)
        {
            case ".tsv":
                return '\t';
            case ".xlsx":
            case ".xls":
            case ".xlsm":
            case ".xlsb":
            case ".ods":
                throw new FormatException(
                    $"Import: '{ext}' files are not yet supported. " +
                    "Save the source as CSV / TSV and try again.");
            default:
                return ',';
        }
    }
}

/// <summary>Validated, fully-resolved instructions for one import
/// operation. The COM-bound service treats this as a command record.</summary>
public sealed record ImportPlan(
    string AbsoluteSourcePath,
    char Delimiter,
    string TargetRangeAddress);
