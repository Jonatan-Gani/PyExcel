using System;
using System.IO;

namespace PyExcel.State;

/// <summary>What a workbook's stored identity says happened to it since it was last
/// committed — derived by comparing the identity carried in its profile against the
/// path it is now open at.</summary>
public enum WorkbookIdentityAction
{
    /// <summary>Same place (or unstamped / no committed origin yet) — nothing to do.</summary>
    Unchanged = 0,

    /// <summary>The file moved or was renamed: its committed origin path no longer
    /// exists, so this is the same project at a new location. Keep the
    /// <see cref="WorkbookProfileData.ProjectId"/>, commit the new origin, and let the
    /// project folder re-resolve (it usually travelled with the workbook).</summary>
    Moved = 1,

    /// <summary>The file is a copy: its committed origin path still exists as a
    /// separate file, so this must become its own project — assign a fresh
    /// <see cref="WorkbookProfileData.ProjectId"/> and detach it from the original's
    /// project folder so the two don't share one environment.</summary>
    Copied = 2,
}

/// <summary>
/// Pure decision logic that reconciles a workbook's <em>stable identity</em>
/// (<see cref="WorkbookProfileData.ProjectId"/> /
/// <see cref="WorkbookProfileData.OriginPath"/>) against the path it is actually open
/// at. This is what lets PyExcel keep following an enabled project across a move or
/// rename, while not letting a Save-As copy silently inherit the original's venv and
/// project folder.
///
/// <para>Kept free of COM and of its own file I/O — the caller supplies the current
/// path and the <c>originExists</c> probe — so the rule is unit-tested on Linux CI;
/// the COM event sink applies the verdict.</para>
/// </summary>
public static class WorkbookIdentityReconciler
{
    /// <summary>Decide what a workbook's identity implies, from the
    /// <paramref name="projectId"/> and <paramref name="originPath"/> read off its
    /// profile, the <paramref name="currentPath"/> it is open at, and whether the
    /// origin still <paramref name="originExists"/> on disk.</summary>
    public static WorkbookIdentityAction Reconcile(
        string? projectId, string? originPath, string? currentPath, bool originExists)
    {
        // Unstamped (a project enabled before identity existed, or a plain workbook):
        // nothing to reconcile until it's stamped on its next save.
        if (string.IsNullOrEmpty(projectId)) return WorkbookIdentityAction.Unchanged;

        // Stamped but with no committed origin, or open with no path yet (unsaved):
        // treat as in-place — the caller commits the current origin.
        if (string.IsNullOrEmpty(originPath) || string.IsNullOrEmpty(currentPath))
            return WorkbookIdentityAction.Unchanged;

        if (PathsEqual(originPath!, currentPath!)) return WorkbookIdentityAction.Unchanged;

        // The path changed: a surviving origin means this is a second copy of the
        // project; a vanished origin means the original moved or was renamed here.
        return originExists ? WorkbookIdentityAction.Copied : WorkbookIdentityAction.Moved;
    }

    /// <summary>Whether two paths denote the same file, by normalised full path,
    /// case-insensitively (the Windows file system). Falls back to an ordinal compare
    /// when a path can't be normalised (e.g. a SharePoint/OneDrive URL).</summary>
    public static bool PathsEqual(string a, string b)
    {
        try
        {
            return string.Equals(
                Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }
    }
}
