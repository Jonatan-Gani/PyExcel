using System;
using System.IO;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>The delimited file format an export produces. Intentionally tiny —
/// the two text formats the v2 export path round-trips with <see cref="CsvWriter"/>
/// / <see cref="CsvParser"/>. Binary spreadsheet formats are out of scope (an
/// export is a flat dump of a range).</summary>
public enum ExportFileType
{
    /// <summary>Comma-separated values (<c>.csv</c>).</summary>
    Csv = 0,
    /// <summary>Tab-separated values (<c>.tsv</c>).</summary>
    Tsv = 1,
}

/// <summary>How (and whether) a date/time stamp is appended to the export file
/// name so each run lands in its own file instead of overwriting the last.
/// The styles are all file-name-safe and culture-invariant.</summary>
public enum ExportTimestampStyle
{
    /// <summary>No stamp — the file name is just the base name (each export
    /// overwrites the previous one of the same name).</summary>
    None = 0,
    /// <summary>Date and time, e.g. <c>2026-06-16_14-30-00</c>.</summary>
    DateAndTime = 1,
    /// <summary>Date only, e.g. <c>2026-06-16</c>.</summary>
    DateOnly = 2,
    /// <summary>Compact date and time, e.g. <c>20260616-143000</c>.</summary>
    Compact = 3,
}

/// <summary>
/// The user-configurable recipe for an export — a source range plus everything
/// needed to name the destination file: the folder to save into, the base file
/// name, the file type, and an optional unique-name date/time stamp. The Edit
/// dialog persists one of these as the workbook's defaults; the Export dialog
/// seeds itself from those defaults, lets the user tweak, and then
/// <see cref="ExportSettingsPlanner"/> resolves it into a concrete
/// <see cref="ExportPlan"/> at run time (stamping the file name with the moment
/// the export actually runs).
///
/// <para>Kept cross-platform (no WinForms, no COM) so the composition rules —
/// stamping, sanitising, extension handling, folder resolution — are unit-tested
/// on Linux CI.</para>
/// </summary>
/// <param name="SourceRange">The range to export (e.g. <c>A1:C10</c> or
/// <c>Sheet1!A1:C10</c>). The one genuinely required field.</param>
/// <param name="Folder">The destination folder. Blank/null means "save next to
/// the workbook"; a relative path resolves against the workbook directory.</param>
/// <param name="BaseName">The file name without an extension or stamp. Blank/null
/// falls back to <see cref="ExportSettingsPlanner.DefaultBaseName"/>.</param>
/// <param name="FileType">The delimited format (and therefore extension) to write.</param>
/// <param name="Timestamp">Whether and how to append a unique date/time stamp.</param>
public sealed record ExportSettings(
    string? SourceRange,
    string? Folder,
    string? BaseName,
    ExportFileType FileType,
    ExportTimestampStyle Timestamp)
{
    /// <summary>An all-default recipe: no range/folder/name yet, CSV, no stamp.</summary>
    public static readonly ExportSettings Defaults =
        new(null, null, null, ExportFileType.Csv, ExportTimestampStyle.None);

    /// <summary>
    /// Build the dialog's seed recipe from a workbook's persisted state. Prefers
    /// the structured export-default fields; when they're empty but a legacy
    /// single-path <see cref="WorkbookState.ExportOutput"/> is set (a workbook
    /// configured before the structured defaults existed), it decomposes that
    /// path into folder / base name / file type so nothing is lost on upgrade.
    /// </summary>
    public static ExportSettings FromState(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        var fileType = ParseFileType(state.ExportFormat);
        var timestamp = ParseTimestamp(state.ExportTimestamp);
        var folder = state.ExportFolder;
        var baseName = state.ExportBaseName;

        // Legacy upgrade: split an older single destination-path field into the
        // structured pieces the new dialog edits.
        if (string.IsNullOrWhiteSpace(folder)
            && string.IsNullOrWhiteSpace(baseName)
            && !string.IsNullOrWhiteSpace(state.ExportOutput))
        {
            var legacy = state.ExportOutput!.Trim();
            var dir = Path.GetDirectoryName(legacy);
            if (!string.IsNullOrEmpty(dir)) folder = dir;
            baseName = Path.GetFileNameWithoutExtension(legacy);
            if (string.Equals(Path.GetExtension(legacy), ".tsv", StringComparison.OrdinalIgnoreCase))
                fileType = ExportFileType.Tsv;
        }

        return new ExportSettings(state.ExportInput, folder, baseName, fileType, timestamp);
    }

    /// <summary>The persisted token for a file type (<c>"csv"</c> / <c>"tsv"</c>).</summary>
    public static string ToToken(ExportFileType type)
        => type == ExportFileType.Tsv ? "tsv" : "csv";

    /// <summary>The persisted token for a timestamp style
    /// (<c>"none"</c> / <c>"datetime"</c> / <c>"date"</c> / <c>"compact"</c>).</summary>
    public static string ToToken(ExportTimestampStyle style) => style switch
    {
        ExportTimestampStyle.DateAndTime => "datetime",
        ExportTimestampStyle.DateOnly => "date",
        ExportTimestampStyle.Compact => "compact",
        _ => "none",
    };

    /// <summary>Parse a persisted file-type token; anything unrecognised (including
    /// null) is <see cref="ExportFileType.Csv"/>, the safe default.</summary>
    public static ExportFileType ParseFileType(string? token)
        => string.Equals(token?.Trim(), "tsv", StringComparison.OrdinalIgnoreCase)
            ? ExportFileType.Tsv
            : ExportFileType.Csv;

    /// <summary>Parse a persisted timestamp-style token; anything unrecognised
    /// (including null) is <see cref="ExportTimestampStyle.None"/>.</summary>
    public static ExportTimestampStyle ParseTimestamp(string? token) => token?.Trim().ToLowerInvariant() switch
    {
        "datetime" => ExportTimestampStyle.DateAndTime,
        "date" => ExportTimestampStyle.DateOnly,
        "compact" => ExportTimestampStyle.Compact,
        _ => ExportTimestampStyle.None,
    };
}

/// <summary>The outcome of the Export dialog's "Export" mode: the recipe to run
/// now, plus whether the user asked to also persist it as the workbook's new
/// export default. A plain data carrier (the WinForms dialog's return type),
/// kept cross-platform here next to <see cref="ExportSettings"/> — mirroring how
/// the other dialogs return cross-platform result types
/// (<see cref="EditIoValidationResult"/>).</summary>
public sealed record ExportPromptResult(ExportSettings Settings, bool SaveAsDefault);
