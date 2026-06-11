using System;
using System.IO;

namespace PyExcel.Excel;

/// <summary>
/// Pure-logic planner for the Phase 5 Import button — validates the
/// user-typed ribbon fields, resolves the source-file path against the
/// active workbook directory, detects the import format from the file
/// extension, and parses the optional <c>path!SheetName</c> syntax used
/// by the Excel-format importer.
///
/// <para>Separated from the COM-bound <c>ImportService</c> so its logic
/// (path resolution, format detection, sheet-syntax parsing, validation
/// messages) is unit-testable on Linux CI without needing Excel.</para>
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
    ///
    /// <para>For Excel-format imports the user can pin a sheet by
    /// appending <c>!SheetName</c> to the source path
    /// (<c>data.xlsx!Q2</c>). Omitted, the importer falls back to the
    /// first sheet.</para>
    /// </summary>
    /// <exception cref="FormatException">A required field is blank, the
    /// file extension names a format we don't support (xls, ods), or a
    /// CSV/TSV path was given with the Excel-only sheet syntax.</exception>
    public static ImportPlan Create(
        string? importInput,
        string? importOutput,
        string? workbookDirectory)
    {
        var source = (importInput ?? string.Empty).Trim();
        if (source.Length == 0)
            throw new FormatException(
                "Import: the Input field is empty. " +
                "Type the path to a CSV / TSV / XLSX file to import from.");

        var target = (importOutput ?? string.Empty).Trim();
        if (target.Length == 0)
            throw new FormatException(
                "Import: the Output field is empty. " +
                "Type the range to write into (e.g. A1, or Sheet1!A1).");

        var (pathPart, sheetName) = ParsePathAndSheet(source);
        var absolutePath = ResolvePath(pathPart, workbookDirectory);
        var format = DetectFormat(absolutePath);

        if (format == ImportFormat.Csv)
        {
            if (sheetName is not null)
                throw new FormatException(
                    "Import: the 'path!Sheet' syntax is only valid for " +
                    "Excel-format files (.xlsx / .xlsm / .xlsb).");
            var delimiter = DetectDelimiter(absolutePath);
            return new ImportPlan(absolutePath, ImportFormat.Csv, delimiter, null, target);
        }

        // Excel: delimiter is unused; sheet name is optional (null = first
        // sheet). The default char ',' is just a placeholder satisfying the
        // record's positional contract.
        return new ImportPlan(absolutePath, ImportFormat.Excel, ',', sheetName, target);
    }

    /// <summary>Split the user-typed input into <c>(path, sheetName?)</c>
    /// using the <c>path!Sheet</c> syntax. Splits on the last <c>!</c>
    /// only when the part before that <c>!</c> ends with an Excel-format
    /// extension (.xlsx / .xlsm / .xlsb) — otherwise the whole input is
    /// the path, so a CSV / TSV with <c>!</c> in its filename is not
    /// mis-parsed. An empty or whitespace-only sheet name is treated as
    /// "no sheet name" (first sheet).</summary>
    public static (string path, string? sheetName) ParsePathAndSheet(string input)
    {
        if (input is null) throw new ArgumentNullException(nameof(input));
        int sep = input.LastIndexOf('!');
        if (sep < 0) return (input, null);

        string left = input.Substring(0, sep);
        string right = input.Substring(sep + 1);
        string ext = Path.GetExtension(left).ToLowerInvariant();
        if (ext is ".xlsx" or ".xlsm" or ".xlsb")
        {
            string sheet = right.Trim();
            return (left, sheet.Length == 0 ? null : sheet);
        }
        return (input, null);
    }

    /// <summary>Compose the user-facing <c>path!Sheet</c> input from a path
    /// and a chosen sheet — the inverse of <see cref="ParsePathAndSheet"/>,
    /// used by the sheet picker to write its choice back into the Import
    /// field. A null/blank sheet returns the path unchanged (first sheet).
    /// Pinning a sheet only round-trips on an Excel-format path, so a sheet
    /// on a non-Excel path is rejected.</summary>
    /// <exception cref="FormatException">A sheet was given for a path that
    /// is not an Excel-format file (.xlsx / .xlsm / .xlsb).</exception>
    public static string Compose(string path, string? sheetName)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));

        var sheet = (sheetName ?? string.Empty).Trim();
        if (sheet.Length == 0) return path;

        var ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext is not (".xlsx" or ".xlsm" or ".xlsb"))
            throw new FormatException(
                "Import: a sheet can only be pinned on an Excel-format path " +
                "(.xlsx / .xlsm / .xlsb).");

        return path + "!" + sheet;
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
    /// <c>.ods</c>) throw <see cref="FormatException"/> — they are not
    /// CSV-shaped and this method is also reused by
    /// <c>ExportPlanner</c>, which still rejects Excel-format
    /// destinations.</summary>
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
                    $"Import: '{ext}' files are not delimited. " +
                    "Use the Excel-format import path instead.");
            default:
                return ',';
        }
    }

    /// <summary>Detect the import format from a file extension. The
    /// modern Excel binaries (<c>.xlsx</c> / <c>.xlsm</c> / <c>.xlsb</c>)
    /// route through the COM-bound Excel importer; everything else
    /// (including no extension, <c>.csv</c>, <c>.tsv</c>, <c>.txt</c>)
    /// falls through to the CSV parser with its own delimiter detection.
    /// <c>.xls</c> (legacy binary) and <c>.ods</c> (OpenDocument) throw
    /// — they're explicitly out of scope for this Phase 5 slice.</summary>
    public static ImportFormat DetectFormat(string path)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".xlsx" or ".xlsm" or ".xlsb" => ImportFormat.Excel,
            ".xls" => throw new FormatException(
                "Import: '.xls' (legacy Excel binary) is not supported. " +
                "Save the source as .xlsx and try again."),
            ".ods" => throw new FormatException(
                "Import: '.ods' (OpenDocument) is not supported. " +
                "Save the source as .xlsx or .csv and try again."),
            _ => ImportFormat.Csv,
        };
    }
}

/// <summary>Which decoder the <c>ImportService</c> should dispatch to.
/// The set is intentionally tiny — Excel binaries on one side, CSV-shaped
/// text on the other. <c>.xls</c> and <c>.ods</c> aren't members because
/// the planner rejects them outright.</summary>
public enum ImportFormat
{
    /// <summary>Comma- or tab-delimited text. Uses <c>CsvParser</c>.</summary>
    Csv = 0,
    /// <summary>Modern Excel binary (.xlsx / .xlsm / .xlsb). Read via COM.</summary>
    Excel = 1,
}

/// <summary>Validated, fully-resolved instructions for one import
/// operation. <see cref="Delimiter"/> is meaningful only when
/// <see cref="Format"/> is <see cref="ImportFormat.Csv"/>;
/// <see cref="SheetName"/> is meaningful only when <see cref="Format"/>
/// is <see cref="ImportFormat.Excel"/> (and may be <see langword="null"/>
/// to mean "first sheet").</summary>
public sealed record ImportPlan(
    string AbsoluteSourcePath,
    ImportFormat Format,
    char Delimiter,
    string? SheetName,
    string TargetRangeAddress);
