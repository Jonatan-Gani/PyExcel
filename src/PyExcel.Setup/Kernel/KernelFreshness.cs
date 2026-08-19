using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.Common.Logging;

namespace PyExcel.Setup.Kernel;

/// <summary>
/// Cached "is the extracted kernel still the one this build ships?" gate.
///
/// <para>Exists because the answer is needed from a ribbon <c>getEnabled</c>
/// callback, which Excel invokes on every ribbon invalidation — several times
/// per selection change. <see cref="KernelResourceExtractor.Check"/> reads and
/// compares every kernel file, which is far too much work for that path, so
/// the verdict is memoised per project directory and only recomputed when
/// something has plausibly changed.</para>
///
/// <para><b>Why this gate matters.</b> The host and the kernel ship in one
/// <c>.xll</c>, but the kernel that actually runs is the copy extracted into
/// the project folder at Setup time. Nothing re-extracts it on its own, so a
/// user who installs a new build keeps running the old kernel — it handshakes
/// cleanly, ignores meta keys it does not understand, and silently behaves like
/// the previous release. That is not hypothetical: a host sending declared
/// input types to a pre-contract kernel got positional dispatch back and a
/// confusing <c>TypeError</c> from inside the user's own script, with nothing
/// anywhere to suggest the two halves were different versions.</para>
/// </summary>
public static class KernelFreshness
{
    /// <summary>Kernel directory name under the project root. Must match
    /// <c>PyExcel.Excel.PythonResolver.ExtractedKernelDirName</c> and the
    /// target <see cref="SetupService"/> extracts into.</summary>
    public const string KernelDirName = ".pyexcel-kernel";

    private static readonly object Gate = new object();
    private static readonly Dictionary<string, bool> Cache =
        new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Whether the kernel extracted under <paramref name="projectDir"/> matches
    /// this build. A missing kernel directory counts as up to date: there is
    /// nothing stale to replace, and offering "Update" on a project that was
    /// never set up would be misleading — that is Enable's job.
    /// </summary>
    public static bool IsUpToDate(string projectDir, ILog? log = null)
    {
        if (string.IsNullOrWhiteSpace(projectDir)) return true;

        lock (Gate)
        {
            if (Cache.TryGetValue(projectDir, out var cached)) return cached;
        }

        var verdict = Probe(projectDir, log);

        lock (Gate)
        {
            Cache[projectDir] = verdict;
        }
        return verdict;
    }

    private static bool Probe(string projectDir, ILog? log)
    {
        var kernelDir = Path.Combine(projectDir, KernelDirName);
        if (!Directory.Exists(kernelDir)) return true;

        try
        {
            var check = new KernelResourceExtractor(log: log).Check(kernelDir);
            if (!check.UpToDate) (log ?? NullLog.Instance).Warn(check.Describe());
            return check.UpToDate;
        }
        catch (Exception ex)
        {
            // Never let a probe failure gate the UI shut. Report and assume
            // fresh — a false "no update available" is recoverable, a ribbon
            // that throws on render is not.
            (log ?? NullLog.Instance).Error(
                $"kernel freshness probe failed for '{kernelDir}'", ex);
            return true;
        }
    }

    /// <summary>
    /// Drop the memoised verdict for <paramref name="projectDir"/>, or for
    /// every project when null. Call after anything that re-extracts the
    /// kernel — Setup, Enable, Update — so the ribbon reflects the new state
    /// instead of the answer from before the run.
    /// </summary>
    public static void Invalidate(string? projectDir = null)
    {
        lock (Gate)
        {
            if (projectDir is null) Cache.Clear();
            else Cache.Remove(projectDir);
        }
    }
}
