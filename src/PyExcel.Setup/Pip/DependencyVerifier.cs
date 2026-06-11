using System;
using System.Collections.Generic;
using System.Linq;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;

namespace PyExcel.Setup.Pip;

/// <summary>
/// Verifies the kernel's required Python packages are importable from
/// the venv produced by <see cref="Python.VenvProvisioner"/>. Runs
/// each module through <c>python -c "import X"</c> and reports a
/// per-package pass/fail.
///
/// <para><b>Why not <c>pip show</c> alone:</b> presence in pip's
/// metadata isn't the same as importability — a corrupted wheel can
/// satisfy <c>pip show</c> and still raise <c>ImportError</c> at
/// runtime. The kernel cares about import success; that's what we
/// check.</para>
///
/// <para><b>No "80% threshold":</b> v1 considered the install
/// successful when most packages imported. v2 is strict — every name
/// in the canonical set must import or the verifier reports failure
/// with the missing list. The wizard can then re-run pip install or
/// surface the underlying error.</para>
/// </summary>
public sealed class DependencyVerifier
{
    /// <summary>
    /// The kernel's canonical import set. Pinned in code (not parsed
    /// from <c>requirements.txt</c>) so the import names and the
    /// distribution names don't have to match — pyarrow's distribution
    /// is "pyarrow" but its import is also "pyarrow", whereas a
    /// distribution like "pillow" imports as "PIL". Today the set
    /// coincides with the requirements file; the indirection keeps us
    /// honest if it ever diverges.
    /// </summary>
    public static readonly IReadOnlyList<string> RequiredModules = new[]
    {
        "pandas",
        "numpy",
        "pyarrow",
        "plotly",
        "matplotlib",
    };

    private readonly ProcessRunner _runner;
    private readonly ILog _log;

    public DependencyVerifier(ProcessRunner? runner = null, ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
        _runner = runner ?? new ProcessRunner(_log);
    }

    /// <summary>
    /// Probe each module in <see cref="RequiredModules"/> and report
    /// per-package status.
    /// </summary>
    public DependencyVerificationResult Verify(string venvPython, int perModuleTimeoutMs = 30_000)
    {
        if (string.IsNullOrWhiteSpace(venvPython))
            throw new ArgumentException("venv python required", nameof(venvPython));

        var statuses = new List<ModuleStatus>();
        foreach (var module in RequiredModules)
        {
            var status = Probe(venvPython, module, perModuleTimeoutMs);
            statuses.Add(status);
            if (status.Importable)
                _log.Info($"  ok: {module}");
            else
                _log.Warn($"  missing: {module} :: {status.FailureReason}");
        }

        var missing = statuses.Where(s => !s.Importable).Select(s => s.Module).ToList();
        var ok = missing.Count == 0;
        if (ok)
            _log.Info($"dependency verification: all {statuses.Count} modules importable");
        else
            _log.Warn(
                $"dependency verification failed: {missing.Count}/{statuses.Count} " +
                $"modules missing ({string.Join(", ", missing)})");

        return new DependencyVerificationResult(ok, statuses);
    }

    private ModuleStatus Probe(string venvPython, string module, int timeoutMs)
    {
        // `-c "import X"` exits 0 silently on success; on failure it
        // emits a Python traceback to stderr with the import error.
        var args = new[] { "-c", $"import {module}" };
        try
        {
            var result = _runner.Run(venvPython, args, timeoutMs: timeoutMs);
            if (result.Success)
                return ModuleStatus.Ok(module);
            return ModuleStatus.Failed(
                module,
                string.IsNullOrWhiteSpace(result.Stderr)
                    ? $"exit {result.ExitCode}"
                    : result.Stderr.Trim());
        }
        catch (Exception ex)
        {
            return ModuleStatus.Failed(
                module, $"{ex.GetType().Name}: {ex.Message}");
        }
    }
}

/// <summary>Per-module verification result.</summary>
public sealed class ModuleStatus
{
    public string Module { get; }
    public bool Importable { get; }
    public string? FailureReason { get; }

    private ModuleStatus(string module, bool importable, string? failureReason)
    {
        Module = module;
        Importable = importable;
        FailureReason = failureReason;
    }

    public static ModuleStatus Ok(string module) => new(module, true, null);
    public static ModuleStatus Failed(string module, string reason) => new(module, false, reason);
}

/// <summary>Aggregated outcome of
/// <see cref="DependencyVerifier.Verify(string, int)"/>.</summary>
public sealed class DependencyVerificationResult
{
    public bool AllImportable { get; }
    public IReadOnlyList<ModuleStatus> Modules { get; }

    public DependencyVerificationResult(bool allImportable, IReadOnlyList<ModuleStatus> modules)
    {
        AllImportable = allImportable;
        Modules = modules;
    }

    public IEnumerable<ModuleStatus> Missing => Modules.Where(m => !m.Importable);
}
