using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace PyExcel.State;

/// <summary>
/// On-disk archive of recent runs: inputs, output, error, plus a
/// human-readable manifest. Each call to <see cref="Archive"/> writes a
/// new directory under <see cref="RootDirectory"/> and then evicts the
/// oldest directories so at most <see cref="MaxRuns"/> remain.
///
/// <para><b>Layout.</b> One directory per run:</para>
/// <code>
/// {root}/
///   20260530T140000123_a1b2c3d4/
///     manifest.txt     # headline metadata, parseable line-per-field
///     input_0.arrow    # one file per positional argument, Arrow IPC bytes
///     input_1.arrow
///     output.arrow     # only when the run produced a payload
///     error.txt        # only on Error / Cancelled, KernelErrorRecord.FormatForClipboard
/// </code>
///
/// <para>The directory name (<c>RunId</c>) embeds the run's start time so
/// lexicographic ordering matches chronological order — pruning then
/// reduces to "keep the last N by name".</para>
///
/// <para><b>Threading.</b> A single coarse lock serialises mutating
/// operations (<see cref="Archive"/>, <see cref="Prune"/>) so concurrent
/// runs don't trample each other's directory creation or pruning. Reads
/// (<see cref="List"/>) take the same lock briefly while enumerating
/// directories — short enough to not be a contention concern given how
/// rarely the UI will list archives.</para>
///
/// <para>Failures inside <see cref="Archive"/> are <i>not</i> swallowed at
/// this layer — the caller (<c>PyRun.Execute*</c>) is in the best
/// position to decide whether an archive write should block the
/// user-facing result. Callers that want best-effort semantics wrap the
/// call in a <c>try/catch</c>.</para>
/// </summary>
public sealed class RunArchive
{
    /// <summary>Default retention cap when none is supplied. Twenty runs
    /// is enough to cover an interactive debugging session without
    /// growing the on-disk footprint to a worrying size.</summary>
    public const int DefaultMaxRuns = 20;

    private readonly object _lock = new();

    /// <summary>Directory the service writes archives under. Created
    /// lazily on the first <see cref="Archive"/> call.</summary>
    public string RootDirectory { get; }

    /// <summary>Maximum number of run directories retained. Older
    /// directories are deleted after each <see cref="Archive"/>.</summary>
    public int MaxRuns { get; }

    public RunArchive(string rootDirectory, int maxRuns = DefaultMaxRuns)
    {
        if (string.IsNullOrWhiteSpace(rootDirectory))
            throw new ArgumentException("rootDirectory must be a non-empty path", nameof(rootDirectory));
        if (maxRuns < 0) throw new ArgumentOutOfRangeException(nameof(maxRuns));
        RootDirectory = rootDirectory;
        MaxRuns = maxRuns;
    }

    /// <summary>
    /// Persist <paramref name="entry"/> as a new directory under
    /// <see cref="RootDirectory"/>; then evict the oldest directories so
    /// at most <see cref="MaxRuns"/> remain.
    /// </summary>
    /// <returns>Absolute path of the run's archive directory.</returns>
    public string Archive(RunArchiveEntry entry)
    {
        if (entry is null) throw new ArgumentNullException(nameof(entry));

        var runId = BuildRunId(entry.Timestamp);
        string dir;
        lock (_lock)
        {
            Directory.CreateDirectory(RootDirectory);
            dir = Path.Combine(RootDirectory, runId);
            Directory.CreateDirectory(dir);

            for (var i = 0; i < entry.Inputs.Count; i++)
            {
                var buffer = entry.Inputs[i];
                if (buffer is null)
                    throw new ArgumentException(
                        $"entry.Inputs[{i}] is null — cannot archive a missing argument buffer",
                        nameof(entry));
                File.WriteAllBytes(Path.Combine(dir, $"input_{i}.arrow"), buffer);
            }

            if (entry.Output is not null)
                File.WriteAllBytes(Path.Combine(dir, "output.arrow"), entry.Output);

            if (entry.Error is not null)
                File.WriteAllText(Path.Combine(dir, "error.txt"), entry.Error.FormatForClipboard());

            File.WriteAllText(Path.Combine(dir, "manifest.txt"), FormatManifest(entry, runId));

            PruneCore();
        }
        return dir;
    }

    /// <summary>
    /// Snapshot the currently-archived runs, newest first. Manifests
    /// that can't be parsed (older schema, corrupted file) are skipped
    /// silently — the archive is best-effort diagnostic data, not load-
    /// bearing storage.
    /// </summary>
    public IReadOnlyList<ArchivedRun> List()
    {
        lock (_lock)
        {
            if (!Directory.Exists(RootDirectory)) return Array.Empty<ArchivedRun>();

            var dirs = Directory.GetDirectories(RootDirectory);
            // Lexicographic descending = chronological descending (newest first)
            // because run IDs lead with yyyyMMddTHHmmssfff.
            Array.Sort(dirs, (a, b) => string.CompareOrdinal(Path.GetFileName(b), Path.GetFileName(a)));

            var result = new List<ArchivedRun>(dirs.Length);
            foreach (var d in dirs)
            {
                if (TryReadArchivedRun(d, out var run)) result.Add(run!);
            }
            return result;
        }
    }

    /// <summary>
    /// Delete the oldest run directories so at most <see cref="MaxRuns"/>
    /// remain. Called automatically at the end of <see cref="Archive"/>;
    /// exposed publicly so tests can drive it deterministically.
    /// </summary>
    public void Prune()
    {
        lock (_lock) { PruneCore(); }
    }

    private void PruneCore()
    {
        if (!Directory.Exists(RootDirectory)) return;
        var dirs = Directory.GetDirectories(RootDirectory);
        if (dirs.Length <= MaxRuns) return;

        // Sort ascending — oldest first — so Skip(MaxRuns) lands on the
        // ones to keep and we delete what comes before.
        Array.Sort(dirs, (a, b) => string.CompareOrdinal(Path.GetFileName(a), Path.GetFileName(b)));
        for (var i = 0; i < dirs.Length - MaxRuns; i++)
        {
            try { Directory.Delete(dirs[i], recursive: true); }
            catch
            {
                // Best-effort — a directory we couldn't delete this round
                // (file lock, transient I/O fault) gets retried by the
                // next Archive call. Worst case: the on-disk footprint
                // briefly exceeds MaxRuns.
            }
        }
    }

    private static string BuildRunId(DateTimeOffset timestamp)
    {
        var ts = timestamp.UtcDateTime.ToString(
            "yyyyMMddTHHmmssfff", CultureInfo.InvariantCulture);
        // Short random suffix so two runs in the same millisecond get
        // distinct directories. Eight hex chars is 32 bits — way more than
        // enough to disambiguate the handful of concurrent runs the
        // supervisor's exchange semaphore actually allows.
        var rand = Guid.NewGuid().ToString("N").Substring(0, 8);
        return $"{ts}_{rand}";
    }

    // -------------------------------------------------------------------------
    // Manifest format
    //
    // One `Key: Value` per line. UTF-8, LF-terminated. Values are
    // single-line — newlines in user-supplied strings (workbook key,
    // script path, error message) are replaced with a literal `\n` so the
    // file stays line-orderly. The full multi-line traceback lives in
    // `error.txt` next door.
    //
    // The order is fixed: a future schema bump renames the file or adds a
    // version line.
    // -------------------------------------------------------------------------

    // Single-quoted T and Z so they're treated as literals — t/T are
    // recognised format specifiers, and even Z (uppercase) has been a
    // source of subtle bugs across .NET versions when left unescaped.
    private const string TimestampFormat = "yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'";

    private static string FormatManifest(RunArchiveEntry entry, string runId)
    {
        var sb = new StringBuilder();
        AppendLine(sb, "RunId", runId);
        AppendLine(sb, "TimestampUtc", entry.Timestamp.UtcDateTime.ToString(
            TimestampFormat, CultureInfo.InvariantCulture));
        AppendLine(sb, "DurationMs", ((long)entry.Duration.TotalMilliseconds)
            .ToString(CultureInfo.InvariantCulture));
        AppendLine(sb, "Source", entry.Source);
        if (entry.WorkbookKey is not null)
            AppendLine(sb, "WorkbookKey", entry.WorkbookKey);
        AppendLine(sb, "ScriptPath", entry.ScriptPath);
        AppendLine(sb, "Function", entry.Function);
        AppendLine(sb, "InputCount", entry.Inputs.Count
            .ToString(CultureInfo.InvariantCulture));
        AppendLine(sb, "Status", entry.Status.ToString());

        if (entry.Error is not null)
        {
            AppendLine(sb, "ErrorCode", entry.Error.Code);
            // Suppress the redundant Type when it equals Code (host-side
            // errors carry the same string in both slots).
            if (!string.Equals(entry.Error.PythonType, entry.Error.Code, StringComparison.Ordinal))
                AppendLine(sb, "ErrorType", entry.Error.PythonType);
            AppendLine(sb, "ErrorMessage", entry.Error.Message);
        }
        return sb.ToString();
    }

    private static void AppendLine(StringBuilder sb, string key, string value)
    {
        sb.Append(key);
        sb.Append(": ");
        // Collapse any embedded line breaks so each field stays on one
        // line. The traceback lives in error.txt; this is just the
        // headline. CRLF is squashed before LF/CR alone so a CRLF doesn't
        // become two spaces.
        sb.Append(value.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' '));
        sb.Append('\n');
    }

    private static bool TryReadArchivedRun(string directory, out ArchivedRun? run)
    {
        run = null;
        var manifestPath = Path.Combine(directory, "manifest.txt");
        if (!File.Exists(manifestPath)) return false;

        Dictionary<string, string> fields;
        try { fields = ParseManifest(File.ReadAllText(manifestPath)); }
        catch { return false; }

        if (!fields.TryGetValue("RunId", out var runId)) return false;
        if (!fields.TryGetValue("TimestampUtc", out var tsStr)) return false;
        if (!fields.TryGetValue("Status", out var statusStr)) return false;
        if (!fields.TryGetValue("ScriptPath", out var script)) return false;
        if (!fields.TryGetValue("Source", out var source)) return false;
        fields.TryGetValue("WorkbookKey", out var workbookKey);

        if (!DateTimeOffset.TryParseExact(
                tsStr, TimestampFormat, CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var ts))
            return false;

        if (!Enum.TryParse<RunArchiveStatus>(statusStr, ignoreCase: false, out var status))
            return false;

        run = new ArchivedRun(
            Directory: directory,
            RunId: runId,
            Timestamp: ts,
            Status: status,
            ScriptPath: script,
            WorkbookKey: workbookKey,
            Source: source);
        return true;
    }

    private static Dictionary<string, string> ParseManifest(string text)
    {
        var fields = new Dictionary<string, string>(StringComparer.Ordinal);
        using var reader = new StringReader(text);
        string? line;
        while ((line = reader.ReadLine()) is not null)
        {
            if (line.Length == 0) continue;
            var sep = line.IndexOf(':');
            if (sep <= 0) continue;
            var key = line.Substring(0, sep);
            var value = line.Length > sep + 1 ? line.Substring(sep + 1).TrimStart(' ') : "";
            fields[key] = value;
        }
        return fields;
    }
}
