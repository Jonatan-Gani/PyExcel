using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Setup.Pip;
using PyExcel.Setup.Python;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="PipRunner"/> tests run real pip commands against a
/// venv created on the fly. We do NOT install anything heavy in
/// these tests — instead we exercise <c>pip show</c> against a
/// package that's already in the venv (pip itself) and use
/// <c>pip install</c> only for a tiny no-op-style dependency
/// when the integration is needed.
/// </summary>
public class PipRunnerTests
{
    [Fact]
    public void Install_NullPaths_Throws()
    {
        var p = new PipRunner();
        Assert.Throws<ArgumentException>(() => p.Install(null!, "/tmp/req.txt"));
        Assert.Throws<ArgumentException>(() => p.Install("py", null!));
    }

    [Fact]
    public void Install_MissingRequirementsFile_Throws()
    {
        var p = new PipRunner();
        Assert.Throws<FileNotFoundException>(() =>
            p.Install("python", "/no/such/requirements.txt"));
    }

    [Fact]
    public void Show_VenvPip_ReportsInstalled()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var project = NewTempDir();
        try
        {
            // A venv always ships pip; querying pip itself proves the
            // PipRunner wires `python -m pip show` correctly without
            // requiring a network install.
            var venv = new VenvProvisioner().Provision(project, python);
            var result = new PipRunner().Show(venv.PythonExecutable, "pip");
            Assert.True(result.Success,
                $"pip show pip exited {result.ExitCode}: {result.Stderr}");
            Assert.Contains("Name: pip", result.Stdout, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Cleanup(project);
        }
    }

    [Fact]
    public void Show_MissingPackage_ReportsFailure()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var project = NewTempDir();
        try
        {
            var venv = new VenvProvisioner().Provision(project, python);
            var result = new PipRunner().Show(venv.PythonExecutable, "this-package-does-not-exist-xyzzy");
            Assert.False(result.Success);
        }
        finally
        {
            Cleanup(project);
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

    private static string NewTempDir()
    {
        var dir = Path.Combine(
            Path.GetTempPath(),
            "pyexcel-pip-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Cleanup(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch { /* best-effort */ }
    }
}
