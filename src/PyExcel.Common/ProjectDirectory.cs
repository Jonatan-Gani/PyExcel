using System;
using System.IO;
using System.Text;

namespace PyExcel.Common;

/// <summary>
/// Resolves the local directory PyExcel uses to host a workbook's per-project
/// Python environment (the <c>.pyexcel-venv</c> and the Setup-extracted
/// <c>.pyexcel-kernel</c>). Both halves must agree on this directory: the
/// Setup wizard provisions into it, and the runtime kernel resolver discovers
/// the venv/kernel from it. Keeping the rule here — referenced by both
/// <c>PyExcel.Setup</c> and <c>PyExcel.Excel</c> — is what stops them drifting.
///
/// <para>Most workbooks live in an ordinary local folder, which is used
/// directly. But a workbook opened from SharePoint or OneDrive-online reports
/// its folder as an <c>https://…</c> URL — there is no local directory to
/// create a venv in, and feeding that URL to <see cref="Path"/> APIs throws
/// "the given path's format is not supported". For those (and any other path
/// that isn't a usable local directory) PyExcel falls back to a stable
/// per-user location under <c>%LOCALAPPDATA%\PyExcel</c>, keyed by a hash of
/// the workbook location so two different cloud workbooks don't collide.</para>
///
/// <para>Resolution precedence:</para>
/// <list type="number">
///   <item>The <c>PYEXCEL_PROJECT_DIR</c> environment override, if set — a
///     single knob to pin every workbook's environment to one local folder
///     (parallels the <c>PYEXCEL_PYTHON</c> interpreter override).</item>
///   <item>The workbook directory itself, when it's a usable local path.</item>
///   <item>The <c>%LOCALAPPDATA%\PyExcel\&lt;hash&gt;</c> fallback, for a
///     cloud/URL/other non-local workbook directory.</item>
/// </list>
///
/// <para>A null/blank workbook directory (an unsaved workbook) with no
/// override is passed through unchanged — the runtime resolver treats that as
/// "no per-project environment, use the bundled defaults", and the Setup
/// wizard never reaches here for an unsaved workbook (the ribbon asks the user
/// to save first).</para>
/// </summary>
public static class ProjectDirectory
{
    /// <summary>Environment variable that pins the project directory for every
    /// workbook, overriding the per-workbook rule.</summary>
    public const string OverrideEnvVar = "PYEXCEL_PROJECT_DIR";

    /// <summary>Resolve the local project directory for a workbook whose folder
    /// is <paramref name="workbookDir"/> (which may be a local path, a
    /// SharePoint/OneDrive URL, or null/blank for an unsaved workbook). The
    /// returned directory is not created — callers create it as needed.</summary>
    public static string? Resolve(string? workbookDir)
    {
        var overridePath = Environment.GetEnvironmentVariable(OverrideEnvVar);
        if (!string.IsNullOrWhiteSpace(overridePath))
        {
            try { return Path.GetFullPath(overridePath!.Trim()); }
            catch { /* malformed override — fall through to the workbook rule */ }
        }

        if (IsUsableLocalPath(workbookDir))
            return Path.GetFullPath(workbookDir!.Trim());

        // An unsaved workbook (no directory at all) has no per-project home —
        // pass it through so the runtime falls back to the bundled defaults.
        if (string.IsNullOrWhiteSpace(workbookDir))
            return workbookDir;

        // A non-local directory (a SharePoint/OneDrive-online URL, etc.):
        // host the environment in a stable per-user local folder instead.
        return Path.Combine(LocalAppDataRoot(), "PyExcel", Hash(workbookDir!.Trim()));
    }

    /// <summary>True when <paramref name="path"/> is a usable local filesystem
    /// directory — non-blank, not a URL, rooted, and accepted by
    /// <see cref="Path.GetFullPath(string)"/>.</summary>
    public static bool IsUsableLocalPath(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return false;
        var p = path!.Trim();
        // A URL (http(s)://server/...) is what Excel hands back for a
        // SharePoint/OneDrive-online workbook — not a local directory.
        if (p.IndexOf("://", StringComparison.Ordinal) >= 0) return false;
        try
        {
            if (!Path.IsPathRooted(p)) return false;
            _ = Path.GetFullPath(p);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string LocalAppDataRoot()
    {
        var dir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        // Some hosts return empty for LocalApplicationData (rare, but defend
        // against it so Combine never throws); fall back to the temp dir.
        return string.IsNullOrEmpty(dir) ? Path.GetTempPath() : dir;
    }

    /// <summary>A stable, process-independent hash of the workbook location,
    /// used as the fallback folder name. FNV-1a over the UTF-8 bytes of the
    /// lower-cased location — deterministic across sessions (unlike
    /// <see cref="string.GetHashCode"/>) so the same cloud workbook always maps
    /// to the same local environment.</summary>
    private static string Hash(string value)
    {
        const ulong offset = 14695981039346656037UL;
        const ulong prime = 1099511628211UL;
        var bytes = Encoding.UTF8.GetBytes(value.ToLowerInvariant());
        ulong hash = offset;
        unchecked
        {
            foreach (var b in bytes)
            {
                hash ^= b;
                hash *= prime;
            }
        }
        return hash.ToString("x16");
    }
}
