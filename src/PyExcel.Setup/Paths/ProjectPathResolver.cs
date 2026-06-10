using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PyExcel.Setup.Paths;

/// <summary>
/// Normalises and classifies a project path so Setup can make the
/// right structural decisions before touching disk. Three path
/// classes Setup distinguishes:
///
/// <list type="number">
///   <item><b>Local</b> — an ordinary drive-letter path
///     (<c>C:\Users\you\PyExcel\Projects\Foo</c>); proceed as usual.</item>
///   <item><b>UNC</b> — a network share (<c>\\server\share\Foo</c> or
///     the long-path <c>\\?\UNC\server\share\Foo</c> form). The path
///     is valid for venv creation only if the share allows executable
///     creation; the resolver flags it so the wizard can warn before
///     attempting <c>python -m venv</c>, which on some SMB shares
///     fails opaquely.</item>
///   <item><b>OneDrive / SharePoint synced</b> — a local mirror of a
///     SharePoint library or OneDrive folder. Detected by matching the
///     path against the <c>OneDrive</c>, <c>OneDriveConsumer</c>, and
///     <c>OneDriveCommercial</c> environment variables that the
///     OneDrive client sets on every Windows session. Marked so the
///     wizard can warn about online-only files and per-folder reserved
///     names a non-default SharePoint library might use.</item>
/// </list>
///
/// <para>The resolver is pure: no disk reads, no env-var writes, no
/// COM calls. It takes a raw path string and returns a
/// <see cref="ProjectPathInfo"/> the caller can act on.</para>
/// </summary>
public sealed class ProjectPathResolver
{
    /// <summary>Resolve and classify <paramref name="rawPath"/>.</summary>
    /// <exception cref="ArgumentException"><paramref name="rawPath"/>
    ///     is null/whitespace, or <see cref="Path.GetFullPath(string)"/>
    ///     rejects it.</exception>
    public ProjectPathInfo Resolve(string rawPath)
    {
        if (string.IsNullOrWhiteSpace(rawPath))
            throw new ArgumentException("path required", nameof(rawPath));

        string normalised;
        try
        {
            normalised = NormaliseUnc(rawPath.Trim());
            normalised = Path.GetFullPath(normalised);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new ArgumentException(
                $"could not normalise path '{rawPath}': {ex.Message}", nameof(rawPath), ex);
        }

        var isUnc = IsUncPath(normalised);
        var (isOneDrive, oneDriveRoot) = ClassifyOneDrive(normalised);

        return new ProjectPathInfo(
            originalPath: rawPath,
            normalisedPath: normalised,
            isUnc: isUnc,
            isOneDriveSynced: isOneDrive,
            oneDriveRoot: oneDriveRoot);
    }

    /// <summary>
    /// Strip the <c>\\?\UNC\</c> long-path prefix that
    /// <see cref="Path.GetFullPath(string)"/> preserves verbatim, so
    /// downstream comparisons and display strings use the canonical
    /// <c>\\server\share\…</c> form.
    /// </summary>
    private static string NormaliseUnc(string path)
    {
        const string prefix = @"\\?\UNC\";
        if (path.StartsWith(prefix, StringComparison.Ordinal))
            return @"\\" + path.Substring(prefix.Length);
        return path;
    }

    private static bool IsUncPath(string normalised)
    {
        // After NormaliseUnc + GetFullPath, a UNC path begins with `\\`
        // (Windows) or contains `//` at the start of the URI-like form.
        // POSIX has no UNC concept — return false there.
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;
        return normalised.StartsWith(@"\\", StringComparison.Ordinal);
    }

    private static (bool IsOneDrive, string? Root) ClassifyOneDrive(string normalised)
    {
        // Iterate every OneDrive root the OS exposes — the personal
        // client, the commercial (SharePoint) client, and the generic
        // alias the original deployment set. A given machine can have
        // any subset of these; OneDriveCommercial is the SharePoint
        // sync client. We pick the longest matching root so a more
        // specific override wins over a parent.
        string?[] candidates =
        {
            Environment.GetEnvironmentVariable("OneDriveCommercial"),
            Environment.GetEnvironmentVariable("OneDriveConsumer"),
            Environment.GetEnvironmentVariable("OneDrive"),
        };

        string? bestMatch = null;
        foreach (var raw in candidates)
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            string root;
            try { root = Path.GetFullPath(raw); }
            catch { continue; }
            if (!IsWithin(normalised, root)) continue;
            if (bestMatch is null || root.Length > bestMatch.Length)
                bestMatch = root;
        }

        return (bestMatch is not null, bestMatch);
    }

    private static bool IsWithin(string path, string parent)
    {
        // Case-insensitive on Windows, case-sensitive elsewhere — match
        // the platform's filesystem semantics. We compare both with a
        // trailing separator so `C:\Foo` doesn't accidentally match
        // `C:\Foobar`.
        var cmp = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        var p = EnsureTrailingSeparator(path);
        var r = EnsureTrailingSeparator(parent);
        return p.StartsWith(r, cmp);
    }

    private static string EnsureTrailingSeparator(string s)
    {
        if (string.IsNullOrEmpty(s)) return s;
        return s[s.Length - 1] == Path.DirectorySeparatorChar
            ? s
            : s + Path.DirectorySeparatorChar;
    }
}

/// <summary>Result of <see cref="ProjectPathResolver.Resolve(string)"/>.</summary>
public sealed class ProjectPathInfo
{
    public string OriginalPath { get; }
    public string NormalisedPath { get; }
    public bool IsUnc { get; }
    public bool IsOneDriveSynced { get; }
    public string? OneDriveRoot { get; }

    public ProjectPathInfo(
        string originalPath,
        string normalisedPath,
        bool isUnc,
        bool isOneDriveSynced,
        string? oneDriveRoot)
    {
        OriginalPath = originalPath;
        NormalisedPath = normalisedPath;
        IsUnc = isUnc;
        IsOneDriveSynced = isOneDriveSynced;
        OneDriveRoot = oneDriveRoot;
    }
}
