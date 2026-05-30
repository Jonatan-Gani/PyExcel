using System;

namespace PyExcel.State;

/// <summary>
/// One row of <see cref="RunArchive.List"/>. The headline fields parsed
/// from <c>manifest.txt</c> so callers can pick a run without re-reading
/// every manifest themselves; <see cref="Directory"/> is the path the
/// caller opens for the full record (inputs, output, error).
/// </summary>
/// <param name="Directory">Absolute path to the run's archive directory.</param>
/// <param name="RunId">Directory leaf — <c>yyyyMMddTHHmmssfff_xxxxxxxx</c>.
/// Sorts lexicographically by chronological order.</param>
/// <param name="Timestamp">When the run started (UTC), parsed from the
/// manifest. Equal to the run-id prefix to millisecond precision.</param>
/// <param name="Status">Outcome of the run, parsed from the manifest.</param>
/// <param name="ScriptPath">Script path, parsed from the manifest.</param>
/// <param name="WorkbookKey">Workbook key that owned the run, or
/// <see langword="null"/> if the run was unbound.</param>
/// <param name="Source">Origin label (<c>"PY.RUN"</c> /
/// <c>"Run Python button"</c>).</param>
/// <param name="HasOutput"><see langword="true"/> iff the archive
/// directory has an <c>output.arrow</c> file — i.e. the run produced a
/// payload the caller can paste back into a range. Distinct from
/// <see cref="Status"/>: a successful run that returned <c>None</c>
/// is <see cref="RunArchiveStatus.Success"/> but
/// <c>HasOutput == false</c>.</param>
public sealed record ArchivedRun(
    string Directory,
    string RunId,
    DateTimeOffset Timestamp,
    RunArchiveStatus Status,
    string ScriptPath,
    string? WorkbookKey,
    string Source,
    bool HasOutput = false);
