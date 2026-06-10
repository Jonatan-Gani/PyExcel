using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Setup.Python;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="SystemPythonProbe"/> tests. The Probe() method that
/// invokes Python relies on a real interpreter being on PATH; CI
/// installs Python 3.12 before running the test suite.
/// </summary>
[Collection("PyExcel.Setup environment")]
public class SystemPythonProbeTests
{
    [Fact]
    public void IsWindowsStoreStub_NonWindowsPath_ReturnsFalse()
    {
        // The stub heuristic is Windows-specific by design — on POSIX
        // there is no WindowsApps directory, so any path must be a
        // real binary as far as the probe is concerned.
        Assert.False(SystemPythonProbe.IsWindowsStoreStub("/usr/bin/python3"));
        Assert.False(SystemPythonProbe.IsWindowsStoreStub("C:\\Python\\python.exe"));
    }

    [Fact]
    public void IsWindowsStoreStub_EmptyPath_ReturnsFalse()
    {
        Assert.False(SystemPythonProbe.IsWindowsStoreStub(""));
        Assert.False(SystemPythonProbe.IsWindowsStoreStub("   "));
        Assert.False(SystemPythonProbe.IsWindowsStoreStub(null!));
    }

    [Fact]
    public void Probe_OverrideMissingFile_ReturnsFailure()
    {
        var prior = Environment.GetEnvironmentVariable(SystemPythonProbe.PythonEnvVar);
        try
        {
            Environment.SetEnvironmentVariable(
                SystemPythonProbe.PythonEnvVar,
                "/this/does/not/exist/python");
            var result = new SystemPythonProbe().Probe();
            Assert.False(result.Found);
            Assert.NotNull(result.FailureReason);
            Assert.Contains("does not point at", result.FailureReason!);
        }
        finally
        {
            Environment.SetEnvironmentVariable(SystemPythonProbe.PythonEnvVar, prior);
        }
    }

    [Fact]
    public void Probe_OverrideValidPython_ReturnsSuccess()
    {
        // CI installs a real Python 3.12. Resolve it via PATH ourselves
        // (mirror what the probe does internally) and feed it back as
        // an explicit override so we exercise both code paths.
        var python = LocatePythonOnPath();
        if (python is null) return; // hosts without Python skip this test silently.

        var prior = Environment.GetEnvironmentVariable(SystemPythonProbe.PythonEnvVar);
        try
        {
            Environment.SetEnvironmentVariable(SystemPythonProbe.PythonEnvVar, python);
            var result = new SystemPythonProbe().Probe();
            Assert.True(result.Found, $"expected found, got: {result.FailureReason}");
            Assert.Equal(python, result.ExecutablePath);
            Assert.NotNull(result.VersionBanner);
            Assert.StartsWith("Python ", result.VersionBanner!);
        }
        finally
        {
            Environment.SetEnvironmentVariable(SystemPythonProbe.PythonEnvVar, prior);
        }
    }

    [Fact]
    public void Probe_PathSearch_FindsRealPython()
    {
        var python = LocatePythonOnPath();
        if (python is null) return;

        var prior = Environment.GetEnvironmentVariable(SystemPythonProbe.PythonEnvVar);
        try
        {
            // Clear the override so the probe falls through to PATH.
            Environment.SetEnvironmentVariable(SystemPythonProbe.PythonEnvVar, null);

            var result = new SystemPythonProbe().Probe();
            Assert.True(result.Found, $"expected found, got: {result.FailureReason}");
            Assert.NotNull(result.ExecutablePath);
        }
        finally
        {
            Environment.SetEnvironmentVariable(SystemPythonProbe.PythonEnvVar, prior);
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
