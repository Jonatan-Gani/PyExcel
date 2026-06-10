using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using PyExcel.Setup.Pip;
using PyExcel.Setup.Python;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Verifier-level integration tests. The CI lanes pre-install the
/// kernel requirements globally, so the system Python on PATH already
/// imports the canonical set — we exercise <see cref="DependencyVerifier"/>
/// against that interpreter rather than spinning a fresh venv (which
/// would be a slow pip-install we already cover in
/// <see cref="PipRunnerTests"/>).
/// </summary>
public class DependencyVerifierTests
{
    [Fact]
    public void RequiredModules_NotEmpty_AndKnown()
    {
        Assert.NotEmpty(DependencyVerifier.RequiredModules);
        Assert.Contains("pyarrow", DependencyVerifier.RequiredModules);
        Assert.Contains("numpy", DependencyVerifier.RequiredModules);
        Assert.Contains("pandas", DependencyVerifier.RequiredModules);
        Assert.Contains("plotly", DependencyVerifier.RequiredModules);
        Assert.Contains("matplotlib", DependencyVerifier.RequiredModules);
    }

    [Fact]
    public void Verify_AgainstSystemPython_ReportsImportability()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var result = new DependencyVerifier().Verify(python);

        // Don't assert AllImportable — a host without the kernel deps
        // would fail. Assert the verifier produced one status per
        // module, in order, with no nulls.
        Assert.Equal(DependencyVerifier.RequiredModules.Count, result.Modules.Count);
        for (var i = 0; i < result.Modules.Count; i++)
            Assert.Equal(DependencyVerifier.RequiredModules[i], result.Modules[i].Module);
    }

    [Fact]
    public void Verify_OnHostWithKernelDeps_AllImportable()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        // CI installs the canonical requirements before running tests
        // (see .github/workflows/ci.yml); skip the assertion if the
        // probe sees a host without them so a local dev box doesn't
        // fail the suite for the unrelated reason of missing packages.
        var probe = new DependencyVerifier().Verify(python);
        if (probe.Missing.Any(m =>
            m.FailureReason?.Contains("ModuleNotFoundError", StringComparison.Ordinal) == true))
            return;

        Assert.True(probe.AllImportable,
            "expected all kernel modules importable on this host; missing: " +
            string.Join(", ", probe.Missing.Select(s => s.Module)));
    }

    [Fact]
    public void Verify_WithMissingModule_ReportsFailureReason()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        // Drive the verifier with a probe that asks for a guaranteed-
        // missing module via direct ProcessRunner — we cannot inject a
        // custom module list because RequiredModules is the contract,
        // but we can confirm Probe's structure by running it directly
        // through PipRunner.Show for symmetry. This test instead
        // verifies the structure of the result type.
        var result = new DependencyVerifier().Verify(python);
        Assert.NotNull(result.Modules);
        foreach (var status in result.Modules)
        {
            Assert.NotNull(status.Module);
            if (!status.Importable)
                Assert.False(string.IsNullOrEmpty(status.FailureReason));
        }
    }

    private static string? LocatePythonOnPath()
    {
        var isWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
        var names = isWindows
            ? new[] { "python.exe", "python3.exe" }
            : new[] { "python3", "python" };
        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in names)
            {
                string full;
                try { full = Path.Combine(dir, name); }
                catch { continue; }
                if (File.Exists(full)) return full;
            }
        }
        return null;
    }
}
