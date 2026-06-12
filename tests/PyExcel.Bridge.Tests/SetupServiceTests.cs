using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using PyExcel.Setup;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// End-to-end <see cref="SetupService"/> integration. One test runs
/// the full pipeline (resolve → ensure → probe → venv → extract →
/// install → verify) against a real Python; the others exercise the
/// failure paths the wizard relies on.
///
/// <para>The full-pipeline test creates a venv and runs pip; it is
/// the slowest test in the suite (tens of seconds on a clean network)
/// and is skipped when Python isn't on PATH or when the
/// <c>PYEXCEL_SETUP_SKIP_PIP</c> environment variable is set to
/// <c>1</c>, which CI can flip for a fast lane if needed.</para>
/// </summary>
[Collection("PyExcel.Setup environment")]
public class SetupServiceTests
{
    [Fact]
    public void Run_NonexistentParentProject_StillSucceedsBecauseDirIsCreated()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;
        if (Environment.GetEnvironmentVariable("PYEXCEL_SETUP_SKIP_PIP") == "1") return;

        // A path that does not yet exist on disk — the ensure-project-dir
        // stage creates it. We deliberately use a deep sub-path to
        // exercise mkdir recursion.
        var target = Path.Combine(
            Path.GetTempPath(),
            "pyexcel-setup-test-" + Guid.NewGuid().ToString("N"),
            "nested", "project");
        try
        {
            var result = new SetupService().Run(target);

            // The stages we expect to see, in order. Even on failure
            // they should be the prefix of this list.
            var expectedOrder = new[]
            {
                "resolve-path",
                "ensure-project-dir",
                "scaffold-project",
                "probe-python",
                "provision-venv",
                "extract-kernel",
                "pip-install",
                "verify-dependencies",
            };

            for (var i = 0; i < result.Steps.Count; i++)
                Assert.Equal(expectedOrder[i], result.Steps[i].Name);

            if (!result.Success)
            {
                // Surface the failing stage for diagnosis in CI output.
                var failed = result.Steps.FirstOrDefault(s => !s.Success);
                Assert.Fail(
                    $"Setup failed at '{failed?.Name}': {failed?.FailureReason}");
            }

            // Side effects we promised: the venv exists, the kernel
            // sources extracted, and the kernel entrypoint is on disk
            // where PyExcel.Excel.PythonResolver looks for it.
            Assert.True(Directory.Exists(Path.Combine(target, ".pyexcel-venv")));
            Assert.True(File.Exists(Path.Combine(target, ".pyexcel-kernel", "pyexcel", "kernel", "__main__.py")));
            // The scaffold stage prepared the user-facing folders too.
            Assert.True(File.Exists(Path.Combine(target, "userScripts", "example.py")));
        }
        finally
        {
            // Clean both the project dir and its synthetic parent.
            var topLevel = Path.GetDirectoryName(Path.GetDirectoryName(target));
            try { if (topLevel is { } && Directory.Exists(topLevel)) Directory.Delete(topLevel, recursive: true); }
            catch { /* best-effort */ }
        }
    }

    [Fact]
    public void Run_EmptyProjectPath_FailsAtResolve()
    {
        var result = new SetupService().Run("");
        Assert.False(result.Success);
        var first = Assert.Single(result.Steps);
        Assert.Equal("resolve-path", first.Name);
        Assert.False(first.Success);
    }

    [Fact]
    public void Run_PythonOverridePointsAtMissingFile_FailsAtProbe()
    {
        var prior = Environment.GetEnvironmentVariable("PYEXCEL_PYTHON");
        try
        {
            Environment.SetEnvironmentVariable(
                "PYEXCEL_PYTHON", "/this/does/not/exist/python");
            var project = NewTempDir();
            try
            {
                var result = new SetupService().Run(project);
                Assert.False(result.Success);
                var probe = result.Steps.FirstOrDefault(s => s.Name == "probe-python");
                Assert.NotNull(probe);
                Assert.False(probe!.Success);
            }
            finally
            {
                Cleanup(project);
            }
        }
        finally
        {
            Environment.SetEnvironmentVariable("PYEXCEL_PYTHON", prior);
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
            "pyexcel-setup-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Cleanup(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch { /* best-effort */ }
    }
}
