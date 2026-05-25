using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PyExcel.Excel;

/// <summary>
/// Locates the Python interpreter and the embedded kernel package for
/// <see cref="KernelHost"/> to spawn. Three-tier resolution:
///
/// <list type="number">
///   <item><c>PYEXCEL_PYTHON</c> environment variable — explicit escape hatch
///     used by tests and power users.</item>
///   <item>A per-workbook venv at <c>&lt;workbook-dir&gt;/.pyexcel-venv/</c>
///     — the production target once <c>PyExcel.Setup</c> (Phase 7) lands.
///     Stubbed for now; the caller passes <see langword="null"/> for the
///     workbook directory until State is in place.</item>
///   <item>PATH-discovered <c>python.exe</c> / <c>python3</c> — fallback
///     so a freshly-installed PyExcel works against the user's existing
///     Python install before Setup runs.</item>
///  </list>
///
/// The embedded-path resolver walks up from the assembly directory looking
/// for a sibling <c>embedded/pyexcel/kernel/__main__.py</c> — same logic
/// the integration tests already use, lifted here so production code
/// uses one source of truth.
/// </summary>
public static class PythonResolver
{
    /// <summary>Environment variable a user (or test) can set to point at
    /// a specific python executable.</summary>
    public const string PythonEnvVar = "PYEXCEL_PYTHON";

    /// <summary>
    /// Resolve a python executable path. Throws
    /// <see cref="FileNotFoundException"/> if nothing usable is found —
    /// callers should let that surface to the user with the configured
    /// search order in the message.
    /// </summary>
    /// <param name="workbookDir">Optional workbook directory; if supplied,
    /// the venv path under <c>.pyexcel-venv</c> is checked before PATH.
    /// Phase 4 callers pass <see langword="null"/>; Phase 3 wires the
    /// actual workbook directory through.</param>
    public static string ResolvePython(string? workbookDir = null)
    {
        var envOverride = Environment.GetEnvironmentVariable(PythonEnvVar);
        if (!string.IsNullOrWhiteSpace(envOverride) && File.Exists(envOverride))
            return envOverride!;

        if (!string.IsNullOrWhiteSpace(workbookDir))
        {
            var venvPython = VenvPythonPath(workbookDir!);
            if (File.Exists(venvPython)) return venvPython;
        }

        var fromPath = SearchPath(WindowsCandidates(), PosixCandidates());
        if (fromPath is { }) return fromPath;

        throw new FileNotFoundException(
            $"could not locate a python executable. " +
            $"Set {PythonEnvVar} to an absolute path, create a .pyexcel-venv " +
            $"next to the workbook (once PyExcel.Setup is in place), or ensure " +
            $"python is on PATH.");
    }

    /// <summary>
    /// Resolve the path to the embedded <c>pyexcel</c> package — i.e. the
    /// directory the kernel needs on <c>PYTHONPATH</c>.
    /// </summary>
    /// <remarks>
    /// Walks up from the calling assembly's directory until it finds a
    /// sibling <c>embedded/pyexcel/kernel/__main__.py</c>, the canonical
    /// marker. In production the .xll ships with <c>embedded/</c> alongside
    /// it; in tests the marker lives at the repo root.
    /// </remarks>
    public static string ResolveEmbeddedPath()
    {
        var startDir = new DirectoryInfo(AppContext.BaseDirectory);
        var found = WalkUpForEmbedded(startDir);
        if (found is { }) return found;

        throw new DirectoryNotFoundException(
            $"could not locate embedded/pyexcel/kernel/__main__.py walking up " +
            $"from {AppContext.BaseDirectory}. The .xll must ship with embedded/ " +
            $"as a sibling directory.");
    }

    // -------------------------------------------------------------------------
    // Internals
    // -------------------------------------------------------------------------

    private static string VenvPythonPath(string workbookDir)
    {
        var venv = Path.Combine(workbookDir, ".pyexcel-venv");
        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? Path.Combine(venv, "Scripts", "python.exe")
            : Path.Combine(venv, "bin", "python");
    }

    private static string[] WindowsCandidates() => new[] { "python.exe", "python3.exe" };
    private static string[] PosixCandidates() => new[] { "python3", "python" };

    private static string? SearchPath(string[] windowsNames, string[] posixNames)
    {
        var names = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? windowsNames
            : posixNames;
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in names)
            {
                var full = Path.Combine(dir, name);
                if (File.Exists(full)) return full;
            }
        }
        return null;
    }

    private static string? WalkUpForEmbedded(DirectoryInfo? start)
    {
        for (var i = 0; i < 8 && start != null; i++)
        {
            var candidate = Path.Combine(
                start.FullName, "embedded", "pyexcel", "kernel", "__main__.py");
            if (File.Exists(candidate))
                return Path.Combine(start.FullName, "embedded");
            start = start.Parent;
        }
        return null;
    }
}
