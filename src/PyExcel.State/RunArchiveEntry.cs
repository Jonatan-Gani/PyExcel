using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// One run, packaged up for <see cref="RunArchive.Archive"/>. Immutable.
/// </summary>
/// <param name="Timestamp">When the run started (UTC). Used to build the
/// archive directory name, so two runs in the same millisecond get
/// distinct directories via a random suffix; lexicographic ordering of
/// directory names matches chronological order.</param>
/// <param name="WorkbookKey">Workbook key of the active workbook at the
/// time of the run, or <see langword="null"/> if no workbook was bound
/// (kernel boot, init).</param>
/// <param name="ScriptPath">Absolute path to the script that ran.</param>
/// <param name="Function">Function name inside the script (default
/// <c>"transform"</c>).</param>
/// <param name="Source">Human-readable origin label, e.g. <c>"PY.RUN"</c>
/// or <c>"Run Python button"</c>. Mirrors <c>KernelErrorRecord.Source</c>.</param>
/// <param name="Duration">Wall-clock duration of the run, including
/// encoding/decoding overhead.</param>
/// <param name="Inputs">Arrow IPC payloads of each positional argument in
/// wire order. May be empty for a no-arg call. Captured raw so a replay
/// can hand them straight back to the kernel.</param>
/// <param name="Output">Arrow IPC payload of the kernel's reply, or
/// <see langword="null"/> if the function returned <c>None</c>, errored,
/// or was cancelled.</param>
/// <param name="Error">Captured error record on <see cref="RunArchiveStatus.Error"/>
/// runs; <see langword="null"/> on success or cancellation.</param>
/// <param name="Status">Outcome of the run.</param>
public sealed record RunArchiveEntry(
    DateTimeOffset Timestamp,
    string? WorkbookKey,
    string ScriptPath,
    string Function,
    string Source,
    TimeSpan Duration,
    IReadOnlyList<byte[]> Inputs,
    byte[]? Output,
    KernelErrorRecord? Error,
    RunArchiveStatus Status);
