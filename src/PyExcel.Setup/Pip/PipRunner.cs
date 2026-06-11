using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;

namespace PyExcel.Setup.Pip;

/// <summary>
/// Invokes <c>pip</c> inside a project venv. Every call goes through
/// <see cref="ProcessRunner"/>, so stdout and stderr land in the
/// PyExcel log file line-by-line as pip emits them — the operator can
/// tail the log during a slow install and watch wheels download.
///
/// <para>Three operations Setup needs:</para>
/// <list type="number">
///   <item><see cref="Install(string, string, int)"/> — bulk install
///     from a requirements file (Setup uses the embedded
///     <c>pyexcel/requirements.txt</c>).</item>
///   <item><see cref="UpgradePip(string, int)"/> — run before
///     <see cref="Install(string, string, int)"/> so a stale
///     bootstrap pip can't fail dependency resolution. Best-effort; a
///     non-zero exit is logged as a warning but does not abort Setup.</item>
///   <item><see cref="Show(string, string, int)"/> — query a specific
///     installed package; the verifier uses this to confirm presence
///     post-install without parsing freeze output.</item>
/// </list>
///
/// <para>Pip is invoked via <c>python -m pip</c>, never the bare
/// <c>pip</c> shim. The shim's location on Windows depends on whether
/// the venv was activated; the module form is unambiguous.</para>
/// </summary>
public sealed class PipRunner
{
    private readonly ProcessRunner _runner;
    private readonly ILog _log;

    public PipRunner(ProcessRunner? runner = null, ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
        _runner = runner ?? new ProcessRunner(_log);
    }

    /// <summary>
    /// Install everything in <paramref name="requirementsPath"/> into
    /// the venv whose python lives at <paramref name="venvPython"/>.
    /// </summary>
    /// <param name="venvPython">Absolute path to the venv's python
    ///     executable (the value <see cref="Python.VenvProvisioner"/>
    ///     returns).</param>
    /// <param name="requirementsPath">Absolute path to a pip-format
    ///     requirements file.</param>
    /// <param name="timeoutMs">Wall-clock timeout; default 15 minutes
    ///     accommodates a fresh wheel build on a slow network.</param>
    /// <returns>Captured exit code + streams. Callers should branch on
    ///     <see cref="ProcessRunResult.Success"/>; a non-zero exit
    ///     means the install failed and <see cref="ProcessRunResult.Stderr"/>
    ///     usually names the offending package.</returns>
    /// <exception cref="ArgumentException">either path is missing.</exception>
    /// <exception cref="FileNotFoundException">the requirements file
    ///     does not exist on disk.</exception>
    public ProcessRunResult Install(
        string venvPython,
        string requirementsPath,
        int timeoutMs = 900_000)
    {
        RequirePath(venvPython, nameof(venvPython));
        RequirePath(requirementsPath, nameof(requirementsPath));
        if (!File.Exists(requirementsPath))
            throw new FileNotFoundException(
                "requirements file not found", requirementsPath);

        var args = new List<string>
        {
            "-m", "pip", "install",
            "--disable-pip-version-check",
            "--no-input",
            "-r", requirementsPath,
        };

        var result = _runner.Run(venvPython, args, timeoutMs: timeoutMs);
        if (!result.Success)
            _log.Warn(
                $"pip install -r {requirementsPath} exited {result.ExitCode}");
        return result;
    }

    /// <summary>
    /// Upgrade pip itself inside the venv. Best-effort:
    /// a non-zero exit is logged but not thrown — Setup continues to
    /// the bulk install and lets that produce the real failure
    /// signal if there is one.
    /// </summary>
    public ProcessRunResult UpgradePip(string venvPython, int timeoutMs = 180_000)
    {
        RequirePath(venvPython, nameof(venvPython));

        var args = new[]
        {
            "-m", "pip", "install",
            "--disable-pip-version-check",
            "--no-input",
            "--upgrade", "pip",
        };

        var result = _runner.Run(venvPython, args, timeoutMs: timeoutMs);
        if (!result.Success)
            _log.Warn(
                $"pip self-upgrade exited {result.ExitCode}; continuing with " +
                $"bundled pip");
        return result;
    }

    /// <summary>
    /// Run <c>pip show &lt;package&gt;</c>. Exit code 0 + non-empty
    /// stdout means installed; non-zero means missing.
    /// </summary>
    public ProcessRunResult Show(string venvPython, string package, int timeoutMs = 30_000)
    {
        RequirePath(venvPython, nameof(venvPython));
        if (string.IsNullOrWhiteSpace(package))
            throw new ArgumentException("package name required", nameof(package));

        var args = new[]
        {
            "-m", "pip", "show",
            "--disable-pip-version-check",
            package,
        };
        return _runner.Run(venvPython, args, timeoutMs: timeoutMs);
    }

    private static void RequirePath(string value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new ArgumentException("path required", name);
    }
}
