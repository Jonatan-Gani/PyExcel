using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;
using PyExcel.Setup.Kernel;
using PyExcel.Setup.Paths;
using PyExcel.Setup.Pip;
using PyExcel.Setup.Python;

namespace PyExcel.Setup;

/// <summary>
/// Top-level facade orchestrating the headless half of Phase 7 — the
/// pieces a UI-less caller (CI, a test, the Phase 8 wizard once it
/// exists) can drive end-to-end:
///
/// <list type="number">
///   <item>Resolve + classify the project path.</item>
///   <item>Ensure the project directory exists.</item>
///   <item>Probe for a usable system Python.</item>
///   <item>Provision the per-project venv at <c>.pyexcel-venv</c>.</item>
///   <item>Extract the embedded kernel sources into
///     <c>.pyexcel-kernel</c> next to the venv.</item>
///   <item>Upgrade pip, then install the canonical requirements file
///     into the venv.</item>
///   <item>Verify every kernel-required module imports.</item>
/// </list>
///
/// <para>Each stage is a discrete log line plus a typed step entry on
/// the returned <see cref="SetupResult"/>. Stage failure short-circuits
/// the remaining stages — Setup never installs against a missing venv
/// or verifies against a failed install.</para>
///
/// <para>This service deliberately does NOT own UI. The Phase 8
/// <c>PyExcel.Forms</c> wizard wraps this with WinForms; tests drive
/// it directly. Splitting headless logic from the UI is what lets
/// Linux CI cover the venv/pip/verify pipeline without WinForms.</para>
/// </summary>
public sealed class SetupService
{
    private readonly ProjectPathResolver _paths;
    private readonly SystemPythonProbe _pythonProbe;
    private readonly VenvProvisioner _venv;
    private readonly KernelResourceExtractor _kernel;
    private readonly PipRunner _pip;
    private readonly DependencyVerifier _verifier;
    private readonly ILog _log;

    /// <summary>
    /// Construct against a logger; every collaborator gets the same
    /// log so a single Setup run produces one ordered trace.
    /// </summary>
    public SetupService(ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
        var runner = new ProcessRunner(_log);
        _paths = new ProjectPathResolver();
        _pythonProbe = new SystemPythonProbe(runner, _log);
        _venv = new VenvProvisioner(runner, _log);
        _kernel = new KernelResourceExtractor(log: _log);
        _pip = new PipRunner(runner, _log);
        _verifier = new DependencyVerifier(runner, _log);
    }

    /// <summary>
    /// Run every Setup stage in order. Returns a structured result
    /// describing each stage's outcome — never throws on a stage
    /// failure, so the caller can render a complete report even when
    /// an earlier stage broke the chain.
    /// </summary>
    public SetupResult Run(string projectPath)
    {
        var steps = new List<SetupStep>();

        var pathInfo = TryRun(steps, "resolve-path",
            () => _paths.Resolve(projectPath));
        if (pathInfo is null) return new SetupResult(steps, success: false);

        var ensure = TryRun(steps, "ensure-project-dir",
            () => { Directory.CreateDirectory(pathInfo.NormalisedPath); return pathInfo.NormalisedPath; });
        if (ensure is null) return new SetupResult(steps, success: false);

        var python = TryRun(steps, "probe-python",
            () =>
            {
                var probe = _pythonProbe.Probe();
                if (!probe.Found)
                    throw new InvalidOperationException(probe.FailureReason ?? "python not found");
                return probe;
            });
        if (python is null) return new SetupResult(steps, success: false);

        var venv = TryRun(steps, "provision-venv",
            () => _venv.Provision(pathInfo.NormalisedPath, python.ExecutablePath!));
        if (venv is null) return new SetupResult(steps, success: false);

        var kernel = TryRun(steps, "extract-kernel",
            () =>
            {
                var target = Path.Combine(pathInfo.NormalisedPath, ".pyexcel-kernel");
                return _kernel.Extract(target);
            });
        if (kernel is null) return new SetupResult(steps, success: false);

        // Requirements ships as a sibling-of-pyexcel resource so it
        // doesn't pollute the importable package. The extractor writes
        // it at <target>/pyexcel/requirements.txt (logical-name path
        // preserved); we hand that to pip directly.
        var requirementsPath = Path.Combine(kernel.TargetDir, "pyexcel", "requirements.txt");

        // pip self-upgrade is best-effort: a non-zero exit is logged
        // inside PipRunner but does not contribute a step entry. If pip
        // itself is broken the next stage will surface a real failure.
        try { _pip.UpgradePip(venv.PythonExecutable); }
        catch (Exception ex) { _log.Warn($"pip self-upgrade skipped: {ex.Message}"); }

        var install = TryRun(steps, "pip-install",
            () =>
            {
                var r = _pip.Install(venv.PythonExecutable, requirementsPath);
                if (!r.Success)
                    throw new InvalidOperationException(
                        $"pip install exited {r.ExitCode}: " +
                        $"{(string.IsNullOrWhiteSpace(r.Stderr) ? r.Stdout : r.Stderr).Trim()}");
                return r;
            });
        if (install is null) return new SetupResult(steps, success: false);

        var verify = TryRun(steps, "verify-dependencies",
            () =>
            {
                var v = _verifier.Verify(venv.PythonExecutable);
                if (!v.AllImportable)
                {
                    var missing = string.Join(", ", v.Missing.Select(m => m.Module));
                    throw new InvalidOperationException(
                        $"dependency verification failed: missing {missing}");
                }
                return v;
            });
        if (verify is null) return new SetupResult(steps, success: false);

        return new SetupResult(steps, success: true);
    }

    private T? TryRun<T>(List<SetupStep> steps, string name, Func<T> body) where T : class
    {
        try
        {
            _log.Info($"setup: {name}");
            var value = body();
            steps.Add(SetupStep.Ok(name));
            return value;
        }
        catch (Exception ex)
        {
            _log.Error($"setup: {name} failed", ex);
            steps.Add(SetupStep.Failed(name, ex.Message));
            return null;
        }
    }
}

/// <summary>One stage's outcome inside a Setup run.</summary>
public sealed class SetupStep
{
    public string Name { get; }
    public bool Success { get; }
    public string? FailureReason { get; }

    private SetupStep(string name, bool success, string? failureReason)
    {
        Name = name;
        Success = success;
        FailureReason = failureReason;
    }

    public static SetupStep Ok(string name) => new(name, true, null);
    public static SetupStep Failed(string name, string reason) => new(name, false, reason);
}

/// <summary>Aggregated outcome of <see cref="SetupService.Run(string)"/>.</summary>
public sealed class SetupResult
{
    public IReadOnlyList<SetupStep> Steps { get; }
    public bool Success { get; }

    public SetupResult(IReadOnlyList<SetupStep> steps, bool success)
    {
        Steps = steps;
        Success = success;
    }
}
