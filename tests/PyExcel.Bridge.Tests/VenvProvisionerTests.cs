using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Setup.Python;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Integration tests for <see cref="VenvProvisioner"/>. Each test
/// creates a real venv in a temp directory and asserts the produced
/// layout matches what <c>PyExcel.Excel.PythonResolver</c> looks for.
/// Skipped silently when no Python is on PATH so the test suite
/// stays green on hosts that lack a Python install.
/// </summary>
public class VenvProvisionerTests
{
    [Fact]
    public void Provision_NullArgs_Throws()
    {
        var v = new VenvProvisioner();
        Assert.Throws<ArgumentException>(() => v.Provision(null!, "py"));
        Assert.Throws<ArgumentException>(() => v.Provision("dir", null!));
    }

    [Fact]
    public void Provision_CreatesVenvAndReportsLayout()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var project = NewTempDir();
        try
        {
            var result = new VenvProvisioner().Provision(project, python);

            Assert.Equal(VenvProvisionOutcome.Created, result.Outcome);
            Assert.True(Directory.Exists(result.VenvDirectory));
            Assert.True(File.Exists(result.PythonExecutable),
                $"venv python missing at {result.PythonExecutable}");

            // The exact filename is OS-dependent — match the resolver's
            // convention rather than asserting an absolute path.
            var expectedName = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                ? "python.exe"
                : "python";
            Assert.Equal(expectedName, Path.GetFileName(result.PythonExecutable));
        }
        finally
        {
            Cleanup(project);
        }
    }

    [Fact]
    public void Provision_TwiceWithSameDir_SecondReturnsAlreadyExists()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var project = NewTempDir();
        try
        {
            var v = new VenvProvisioner();
            v.Provision(project, python);
            var second = v.Provision(project, python);
            Assert.Equal(VenvProvisionOutcome.AlreadyExists, second.Outcome);
        }
        finally
        {
            Cleanup(project);
        }
    }

    [Fact]
    public void VenvPythonPath_ReturnsResolverCompatibleLayout()
    {
        var path = VenvProvisioner.VenvPythonPath("/tmp/proj/.pyexcel-venv");
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Assert.EndsWith(Path.Combine("Scripts", "python.exe"), path);
        }
        else
        {
            Assert.EndsWith(Path.Combine("bin", "python"), path);
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
            "pyexcel-venv-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Cleanup(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch { /* best-effort */ }
    }
}
