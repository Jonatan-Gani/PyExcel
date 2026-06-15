using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Pure-logic planner for picking which archived Python run to paste back —
/// validates the user-typed target range and selects the run whose
/// <c>output.arrow</c> would be decoded and written into that range.
///
/// <para><b>Currently unused by the ribbon.</b> The Paste button is now a plain
/// OS-clipboard paste (<see cref="PasteService"/>), so nothing wires this planner
/// in at present. It's kept — and still unit-tested by <c>PastePlannerTests</c> —
/// because the "drop the most recent run's archived output into a range" capability
/// it encodes may come back as its own affordance. Delete it (and its tests) if
/// that's ruled out.</para>
///
/// <para><b>Selection rule.</b> The newest archived run that
/// <em>produced output</em> wins (<see cref="ArchivedRun.HasOutput"/>),
/// regardless of <see cref="ArchivedRun.Status"/> — a script that failed
/// but already emitted partial output (rare but possible) is still
/// pasteable, and the file existence is the most reliable signal. When a
/// <paramref name="workbookKey"/> is supplied, the selection filters to
/// runs bound to that workbook; if no such run exists, the planner falls
/// back to unbound (workbook-less) runs so a fresh workbook can still
/// paste from a recent <c>=PY.RUN</c> made in an unbound context.</para>
///
/// <para>Cross-platform: takes an already-materialised list of runs
/// instead of a <see cref="RunArchive"/> instance so the planner stays
/// pure (no I/O, no statics) and Linux CI can drive it directly.</para>
/// </summary>
public static class PastePlanner
{
    /// <summary>
    /// Build a paste plan from the user's ribbon state and a snapshot of
    /// recent archived runs (typically <see cref="RunArchive.List"/>).
    /// </summary>
    /// <param name="pasteOutput">Raw text from the Paste-Output ribbon
    /// field — the target range address.</param>
    /// <param name="workbookKey">The active workbook's key, used to
    /// prefer same-workbook runs over unbound ones. Pass
    /// <see langword="null"/> when no workbook is active to skip the
    /// preference and just take the newest output-bearing run.</param>
    /// <param name="recentRuns">Archived runs newest first. Empty is
    /// fine — the planner throws <see cref="FormatException"/> with a
    /// clear message instead of returning null.</param>
    /// <exception cref="FormatException"><paramref name="pasteOutput"/>
    /// is blank, or <paramref name="recentRuns"/> contains no
    /// output-bearing run.</exception>
    public static PastePlan Create(
        string? pasteOutput,
        string? workbookKey,
        IReadOnlyList<ArchivedRun> recentRuns)
    {
        if (recentRuns is null) throw new ArgumentNullException(nameof(recentRuns));

        var target = (pasteOutput ?? string.Empty).Trim();
        if (target.Length == 0)
            throw new FormatException(
                "Paste: the Output field is empty. " +
                "Type the range to paste into (e.g. A1, or Sheet1!A1).");

        var selected = SelectRun(workbookKey, recentRuns);
        if (selected is null)
            throw new FormatException(
                "Paste: no recent run has produced output. " +
                "Run a script that returns a value, then try again.");

        var outputPath = Path.Combine(selected.Directory, "output.arrow");
        return new PastePlan(outputPath, target, selected.RunId);
    }

    /// <summary>Pick the newest output-bearing run, preferring runs
    /// bound to <paramref name="workbookKey"/> over unbound ones. Lists
    /// are assumed to be newest-first per <see cref="RunArchive.List"/>'s
    /// contract; we walk linearly and pick the first match.</summary>
    private static ArchivedRun? SelectRun(string? workbookKey, IReadOnlyList<ArchivedRun> recentRuns)
    {
        if (workbookKey is not null)
        {
            // First pass — same-workbook preference.
            for (int i = 0; i < recentRuns.Count; i++)
            {
                var r = recentRuns[i];
                if (r.HasOutput
                    && string.Equals(r.WorkbookKey, workbookKey, StringComparison.Ordinal))
                    return r;
            }
        }

        // Fallback — any output-bearing run (regardless of binding).
        for (int i = 0; i < recentRuns.Count; i++)
        {
            var r = recentRuns[i];
            if (r.HasOutput) return r;
        }
        return null;
    }
}

/// <summary>Validated, fully-resolved instructions for one paste
/// operation. <see cref="SourceRunId"/> is informational — surfaced to
/// the user in LogDisplay so they can correlate the paste against the
/// archive when debugging.</summary>
public sealed record PastePlan(
    string SourceArrowPath,
    string TargetRangeAddress,
    string SourceRunId);
