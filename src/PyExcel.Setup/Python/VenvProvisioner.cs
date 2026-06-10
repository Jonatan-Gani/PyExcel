using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;

namespace PyExcel.Setup.Python;

/// <summary>
/// Creates a per-project Python virtual environment at
/// <c>&lt;projectDir&gt;/.pyexcel-venv</c> — the location
/// <c>PyExcel.Excel.PythonResolver.ResolvePython</c> already looks for
/// at runtime. The shape (Scripts\python.exe on Windows, bin/python on
/// POSIX) matches the resolver verbatim so a venv created here is
/// immediately discoverable.
///
/// <para>Idempotent: an existing, working venv is left untouched
/// (<see cref="Provision"/> returns <see cref="VenvProvisionOutcome.AlreadyExists"/>).
/// A directory that exists but is missing the executable is treated as
/// corruption and recreated from scratch — the alternative is silently
/// running pip against a half-built environment, which produces
/// confusing errors three steps later.</para>
///
/// <para>The created venv inherits the system Python's site packages
/// by default in v1; v2 deliberately does NOT inherit them
/// (<c>--system-site-packages</c> is not passed), so a transform script's
/// dependencies are reproducible across machines.</para>
/// </summary>
public sealed class VenvProvisioner
{
    /// <summary>Name of the venv directory PyExcel creates next to the
    /// project. Matches <c>PyExcel.Excel.PythonResolver</c>'s lookup.</summary>
    public const string VenvDirectoryName = ".pyexcel-venv";

    private readonly ProcessRunner _runner;
    private readonly ILog _log;

    public VenvProvisioner(ProcessRunner? runner = null, ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
        _runner = runner ?? new ProcessRunner(_log);
    }

    /// <summary>
    /// Create (or reuse) a venv at <c>projectDir/.pyexcel-venv</c>.
    /// </summary>
    /// <param name="projectDir">Absolute path to the project directory.
    ///     Created if missing.</param>
    /// <param name="systemPythonPath">Absolute path to the interpreter
    ///     <see cref="SystemPythonProbe"/> resolved.</param>
    /// <returns>The venv location and the python executable path
    ///     inside it; identical shape regardless of whether the venv
    ///     was just created or already existed.</returns>
    /// <exception cref="ArgumentException">either path argument is
    ///     null/whitespace.</exception>
    /// <exception cref="InvalidOperationException">the venv command
    ///     failed; the message carries the captured stderr.</exception>
    public VenvProvisionResult Provision(string projectDir, string systemPythonPath)
    {
        if (string.IsNullOrWhiteSpace(projectDir))
            throw new ArgumentException("project dir required", nameof(projectDir));
        if (string.IsNullOrWhiteSpace(systemPythonPath))
            throw new ArgumentException("system python path required", nameof(systemPythonPath));

        Directory.CreateDirectory(projectDir);
        var venvDir = Path.Combine(projectDir, VenvDirectoryName);
        var venvPython = VenvPythonPath(venvDir);

        if (Directory.Exists(venvDir) && File.Exists(venvPython))
        {
            _log.Info($"venv already present at {venvDir}");
            return new VenvProvisionResult(
                venvDir, venvPython, VenvProvisionOutcome.AlreadyExists);
        }

        if (Directory.Exists(venvDir))
        {
            _log.Warn(
                $"venv directory {venvDir} exists but is missing the python " +
                $"executable; recreating");
            TryDelete(venvDir);
        }

        // `python -m venv` accepts the absolute target dir. We pass it
        // unmodified so a UNC or spaces-in-path project still works —
        // ProcessRunner quotes the argv.
        var result = _runner.Run(
            systemPythonPath,
            new[] { "-m", "venv", venvDir },
            timeoutMs: 120_000);

        if (!result.Success)
            throw new InvalidOperationException(
                $"venv creation failed (exit {result.ExitCode}): " +
                $"{Trim(result.Stderr)}");

        if (!File.Exists(venvPython))
            throw new InvalidOperationException(
                $"venv command succeeded but {venvPython} was not produced; " +
                $"the system python at {systemPythonPath} may be missing the " +
                $"venv module");

        _log.Info($"venv created at {venvDir}");
        return new VenvProvisionResult(
            venvDir, venvPython, VenvProvisionOutcome.Created);
    }

    /// <summary>
    /// Return the path to the python executable inside a venv rooted at
    /// <paramref name="venvDir"/>. Matches the layout used by
    /// <c>PyExcel.Excel.PythonResolver</c>.
    /// </summary>
    public static string VenvPythonPath(string venvDir)
    {
        if (string.IsNullOrWhiteSpace(venvDir))
            throw new ArgumentException("venv dir required", nameof(venvDir));
        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? Path.Combine(venvDir, "Scripts", "python.exe")
            : Path.Combine(venvDir, "bin", "python");
    }

    private static void TryDelete(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"could not delete partial venv at {dir}: {ex.Message}", ex);
        }
    }

    private static string Trim(string s) => string.IsNullOrEmpty(s) ? string.Empty : s.Trim();
}

/// <summary>Outcome of <see cref="VenvProvisioner.Provision"/>.</summary>
public sealed class VenvProvisionResult
{
    public string VenvDirectory { get; }
    public string PythonExecutable { get; }
    public VenvProvisionOutcome Outcome { get; }

    public VenvProvisionResult(string venvDirectory, string pythonExecutable, VenvProvisionOutcome outcome)
    {
        VenvDirectory = venvDirectory;
        PythonExecutable = pythonExecutable;
        Outcome = outcome;
    }
}

public enum VenvProvisionOutcome
{
    Created,
    AlreadyExists,
}
