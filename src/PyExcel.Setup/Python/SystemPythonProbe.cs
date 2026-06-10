using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;

namespace PyExcel.Setup.Python;

/// <summary>
/// Locates a usable system Python interpreter — the one Setup hands to
/// <see cref="VenvProvisioner"/> when creating a per-project venv. The
/// runtime kernel resolver (<c>PyExcel.Excel.PythonResolver</c>) prefers
/// the per-workbook venv over PATH; this probe is the one-shot pre-venv
/// lookup.
///
/// <para>Three failure modes Setup must distinguish so the wizard can
/// produce an actionable message:</para>
/// <list type="number">
///   <item><b>No Python at all</b> — neither <c>PYEXCEL_PYTHON</c> nor
///     PATH yields an executable.</item>
///   <item><b>Windows Store stub</b> — a zero-byte
///     <c>%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe</c> shim that
///     opens the Store page when invoked. PATH ordering on a default
///     Windows install puts this ahead of a real Python, so a naive
///     "is it on PATH" check returns the stub. We detect it by path and
///     by the file's tiny size (Microsoft ships the stub at ~0 bytes
///     plus a reparse point).</item>
///   <item><b>Resolved but broken</b> — the executable exists but
///     <c>python --version</c> fails. The wizard surfaces stdout/stderr
///     verbatim because the cause (corrupted install, missing DLL,
///     permission denied) is almost always there.</item>
/// </list>
/// </summary>
public sealed class SystemPythonProbe
{
    /// <summary>Environment variable mirroring
    /// <c>PyExcel.Excel.PythonResolver.PythonEnvVar</c> — Setup respects
    /// the same explicit override so an operator who has the runtime
    /// resolver pinned doesn't need a second variable.</summary>
    public const string PythonEnvVar = "PYEXCEL_PYTHON";

    private readonly ProcessRunner _runner;
    private readonly ILog _log;

    public SystemPythonProbe(ProcessRunner? runner = null, ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
        _runner = runner ?? new ProcessRunner(_log);
    }

    /// <summary>
    /// Find a Python executable and report its version. Returns a
    /// failure result rather than throwing so the wizard can render the
    /// reason inline — every failure carries the search path it tried.
    /// </summary>
    public PythonProbeResult Probe()
    {
        var envOverride = Environment.GetEnvironmentVariable(PythonEnvVar);
        if (!string.IsNullOrWhiteSpace(envOverride))
        {
            if (!File.Exists(envOverride))
                return PythonProbeResult.Failed(
                    $"{PythonEnvVar}={envOverride} does not point at an existing file");
            return Inspect(envOverride!);
        }

        var (path, stubDetected) = SearchPath();
        if (path is null)
        {
            var message = stubDetected
                ? "the Python on PATH is the Windows Store stub at " +
                  "%LOCALAPPDATA%\\Microsoft\\WindowsApps\\python.exe; install Python " +
                  "from python.org and re-open Excel"
                : "no python executable on PATH; install Python 3 and re-open Excel " +
                  $"or set {PythonEnvVar} to an absolute path";
            return PythonProbeResult.Failed(message);
        }

        return Inspect(path);
    }

    /// <summary>
    /// Is <paramref name="path"/> the Windows Store stub interpreter?
    /// Exposed for tests and for <see cref="VenvProvisioner"/> to
    /// re-check when handed an external value.
    /// </summary>
    public static bool IsWindowsStoreStub(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return false;
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return false;

        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (string.IsNullOrEmpty(localAppData)) return false;
        var stubDir = Path.Combine(localAppData, "Microsoft", "WindowsApps");

        // Normalised case-insensitive prefix check — WindowsApps lives
        // under %LOCALAPPDATA% on every install. The stub binary sits
        // directly inside this folder, not in a sub-tree.
        var normalisedPath = Path.GetFullPath(path);
        if (!normalisedPath.StartsWith(stubDir, StringComparison.OrdinalIgnoreCase))
            return false;

        try
        {
            var info = new FileInfo(normalisedPath);
            // The stub is a reparse point that surfaces as a small file
            // (typically a few hundred bytes); a real python.exe is
            // measured in megabytes. Cap at 1 MiB to leave headroom.
            const long stubByteCap = 1024 * 1024;
            if ((info.Attributes & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint)
                return true;
            if (info.Length > 0 && info.Length < stubByteCap)
                return true;
            return false;
        }
        catch
        {
            // If we can't even stat the file, treat the path as
            // suspicious — Setup should refuse to build a venv on top
            // of something it can't reason about.
            return true;
        }
    }

    // -------------------------------------------------------------------------
    // Internals
    // -------------------------------------------------------------------------

    private PythonProbeResult Inspect(string path)
    {
        if (IsWindowsStoreStub(path))
            return PythonProbeResult.Failed(
                $"the python executable at {path} is the Windows Store stub; install " +
                $"Python from python.org or set {PythonEnvVar} to a real interpreter");

        try
        {
            var result = _runner.Run(path, new[] { "--version" }, timeoutMs: 10_000);
            if (!result.Success)
                return PythonProbeResult.Failed(
                    $"`{path} --version` exited {result.ExitCode}: " +
                    $"{(string.IsNullOrWhiteSpace(result.Stderr) ? result.Stdout : result.Stderr).Trim()}");

            // CPython prints "Python 3.12.4" to stdout on 3.4+; older
            // versions printed it to stderr. Combine before parsing so
            // the probe works against both.
            var combined = (result.Stdout + "\n" + result.Stderr).Trim();
            return PythonProbeResult.Success(path, combined);
        }
        catch (Exception ex)
        {
            return PythonProbeResult.Failed(
                $"failed to invoke {path}: {ex.GetType().Name}: {ex.Message}");
        }
    }

    private static (string? Path, bool StubDetected) SearchPath()
    {
        var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        var names = isWindows
            ? new[] { "python.exe", "python3.exe" }
            : new[] { "python3", "python" };

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        var stubSeen = false;

        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in names)
            {
                string full;
                try { full = Path.Combine(dir, name); }
                catch { continue; }
                if (!File.Exists(full)) continue;

                if (IsWindowsStoreStub(full))
                {
                    stubSeen = true;
                    continue;
                }
                return (full, stubSeen);
            }
        }
        return (null, stubSeen);
    }
}

/// <summary>Outcome of <see cref="SystemPythonProbe.Probe"/>.</summary>
public sealed class PythonProbeResult
{
    public bool Found { get; }
    public string? ExecutablePath { get; }
    public string? VersionBanner { get; }
    public string? FailureReason { get; }

    private PythonProbeResult(bool found, string? path, string? version, string? reason)
    {
        Found = found;
        ExecutablePath = path;
        VersionBanner = version;
        FailureReason = reason;
    }

    public static PythonProbeResult Success(string path, string versionBanner)
        => new(true, path, versionBanner, null);

    public static PythonProbeResult Failed(string reason)
        => new(false, null, null, reason);
}
