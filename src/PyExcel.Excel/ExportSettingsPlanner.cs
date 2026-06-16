using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace PyExcel.Excel;

/// <summary>File-type and timestamp-style helpers — the small, pure mappings
/// from the <see cref="ExportFileType"/> / <see cref="ExportTimestampStyle"/>
/// enums to extensions, delimiters, human labels, and rendered stamps. Kept
/// next to <see cref="ExportSettingsPlanner"/> because that is their only
/// consumer (plus the dialog's drop-downs).</summary>
public static class ExportFormatExtensions
{
    /// <summary>The file extension (with leading dot) for a file type.</summary>
    public static string Extension(this ExportFileType type)
        => type == ExportFileType.Tsv ? ".tsv" : ".csv";

    /// <summary>The field delimiter character for a file type.</summary>
    public static char Delimiter(this ExportFileType type)
        => type == ExportFileType.Tsv ? '\t' : ',';

    /// <summary>A human-readable label for a file type, for the dialog drop-down.</summary>
    public static string Label(this ExportFileType type) => type == ExportFileType.Tsv
        ? "Tab-separated (.tsv)"
        : "Comma-separated (.csv)";

    /// <summary>Render the date/time stamp for a style at the given moment, or the
    /// empty string for <see cref="ExportTimestampStyle.None"/>. All formats are
    /// file-name-safe (no <c>:</c>) and invariant-culture so the same instant
    /// always produces the same stamp regardless of the machine's locale.</summary>
    public static string ToStamp(this ExportTimestampStyle style, DateTime when) => style switch
    {
        ExportTimestampStyle.DateAndTime => when.ToString("yyyy-MM-dd_HH-mm-ss", CultureInfo.InvariantCulture),
        ExportTimestampStyle.DateOnly => when.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        ExportTimestampStyle.Compact => when.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture),
        _ => string.Empty,
    };

    /// <summary>A human-readable example of a timestamp style, for the dialog
    /// drop-down (rendered against a fixed sample instant so the labels are
    /// stable). <see cref="ExportTimestampStyle.None"/> has no example.</summary>
    public static string Example(this ExportTimestampStyle style)
        => style == ExportTimestampStyle.None
            ? string.Empty
            : style.ToStamp(new DateTime(2026, 6, 16, 14, 30, 0));
}

/// <summary>
/// Pure-logic resolver that turns an <see cref="ExportSettings"/> recipe into a
/// concrete <see cref="ExportPlan"/> (source range + absolute target path +
/// delimiter), composing the destination file name from the base name, optional
/// date/time stamp, and file-type extension, and resolving the folder against the
/// workbook directory.
///
/// <para>This is the "smart" core the Export dialog and the run path share, kept
/// cross-platform so the naming rules (stamping, sanitising, extension handling,
/// folder resolution) are unit-tested on Linux CI. The single genuinely required
/// field is the source range; everything else has a sensible default
/// (<see cref="DefaultBaseName"/>, the workbook folder, CSV, no stamp), so a
/// barely-configured export still produces a valid file.</para>
/// </summary>
public static class ExportSettingsPlanner
{
    /// <summary>The base file name used when the user hasn't typed one.</summary>
    public const string DefaultBaseName = "export";

    /// <summary>Characters never allowed in a file name. A fixed set (the Windows
    /// reserved set) rather than <see cref="Path.GetInvalidFileNameChars"/> so the
    /// sanitiser behaves identically on the Windows product and the Linux CI that
    /// tests it — <see cref="Path.GetInvalidFileNameChars"/> returns only
    /// <c>/</c> and NUL on Linux, which would let a typed <c>:</c> or <c>*</c>
    /// through in a test but not in production.</summary>
    private static readonly char[] InvalidNameChars =
        { '<', '>', ':', '"', '/', '\\', '|', '?', '*' };

    /// <summary>
    /// Resolve a recipe into a runnable plan, stamping the file name with
    /// <paramref name="timestamp"/> (pass the moment the export runs so each run
    /// is uniquely named). <paramref name="workbookDirectory"/> is the active
    /// workbook's folder when saved, else null (a blank or relative destination
    /// folder then resolves against the process working directory).
    /// </summary>
    /// <exception cref="FormatException">The source range is blank.</exception>
    public static ExportPlan Resolve(
        ExportSettings settings, DateTime timestamp, string? workbookDirectory)
    {
        if (settings is null) throw new ArgumentNullException(nameof(settings));

        var source = (settings.SourceRange ?? string.Empty).Trim();
        if (source.Length == 0)
            throw new FormatException(
                "Export: choose a source range to export " +
                "(e.g. A1:C10, or Sheet1!A1:C10).");

        var fileName = ComposeFileName(settings, timestamp);
        var folder = (settings.Folder ?? string.Empty).Trim();
        var directory = folder.Length == 0
            ? (string.IsNullOrWhiteSpace(workbookDirectory)
                ? Environment.CurrentDirectory
                : workbookDirectory!)
            : ImportPlanner.ResolvePath(folder, workbookDirectory);

        var absolutePath = Path.GetFullPath(Path.Combine(directory, fileName));
        return new ExportPlan(source, absolutePath, settings.FileType.Delimiter());
    }

    /// <summary>Compose the destination file name (no folder) from the base name,
    /// the optional stamp, and the file-type extension — e.g.
    /// <c>report_2026-06-16_14-30-00.csv</c>. The base name is sanitised and falls
    /// back to <see cref="DefaultBaseName"/> when blank, so this never throws.</summary>
    public static string ComposeFileName(ExportSettings settings, DateTime timestamp)
    {
        if (settings is null) throw new ArgumentNullException(nameof(settings));

        var baseName = SanitizeBaseName(settings.BaseName);
        if (baseName.Length == 0) baseName = DefaultBaseName;

        var stamp = settings.Timestamp.ToStamp(timestamp);
        var name = stamp.Length == 0 ? baseName : baseName + "_" + stamp;
        return name + settings.FileType.Extension();
    }

    /// <summary>A stable, human-readable preview of the file name a recipe
    /// produces, with a literal <c>{timestamp}</c> placeholder in place of the
    /// live stamp — for ribbon/labels that shouldn't tick every second. Never
    /// throws.</summary>
    public static string PreviewPattern(ExportSettings settings)
    {
        if (settings is null) throw new ArgumentNullException(nameof(settings));

        var baseName = SanitizeBaseName(settings.BaseName);
        if (baseName.Length == 0) baseName = DefaultBaseName;

        var name = settings.Timestamp == ExportTimestampStyle.None
            ? baseName
            : baseName + "_{timestamp}";
        return name + settings.FileType.Extension();
    }

    /// <summary>Clean a user-typed base name into something safe to put on disk:
    /// trim it, drop a redundant trailing <c>.csv</c> / <c>.tsv</c> the user may
    /// have typed (so the extension isn't doubled), and strip any file-name-illegal
    /// characters. Returns the empty string when nothing usable remains (the caller
    /// substitutes <see cref="DefaultBaseName"/>).</summary>
    public static string SanitizeBaseName(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;

        var name = raw!.Trim();
        foreach (var ext in new[] { ".csv", ".tsv" })
        {
            if (name.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
            {
                name = name.Substring(0, name.Length - ext.Length);
                break;
            }
        }

        var sb = new StringBuilder(name.Length);
        foreach (var ch in name)
        {
            // Drop control characters and the reserved set; keep everything else.
            if (ch >= ' ' && Array.IndexOf(InvalidNameChars, ch) < 0)
                sb.Append(ch);
        }

        return sb.ToString().Trim();
    }
}
