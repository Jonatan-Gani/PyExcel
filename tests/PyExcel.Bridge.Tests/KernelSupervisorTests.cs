using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Bridge;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// End-to-end test: KernelSupervisor spawns a real
/// <c>python -m pyexcel.kernel</c>, runs the HELLO handshake, exercises
/// PING/PONG, and shuts down cleanly with exit code 0.
///
/// Skipped on Windows for now — the Python transport layer's Win32
/// named-pipe client isn't implemented yet (the kernel currently only has
/// the POSIX/AF_UNIX backend). The Windows CI job still validates the
/// non-integration tests in this assembly.
/// </summary>
public class KernelSupervisorTests
{
    [Fact]
    public void Spawn_Handshake_Ping_Shutdown_RoundTrip()
    {
        if (ShouldSkipForPlatform()) return;

        var python = DiscoverPython();
        var pythonPath = DiscoverEmbeddedPath();
        Assert.False(string.IsNullOrEmpty(python),
            "no python3/python executable found on PATH");
        Assert.False(string.IsNullOrEmpty(pythonPath),
            "embedded/pyexcel/kernel/__main__.py not found relative to test binary");

        using var sup = KernelSupervisor.StartPython(python!, pythonPath!);

        Assert.Equal(Framing.ProtocolVersion, sup.RemoteProtocolVersion);
        Assert.False(sup.Process.HasExited, "kernel exited during handshake");

        var rtt = sup.Ping(timeoutMs: 2000);
        Assert.True(rtt < TimeSpan.FromSeconds(2),
            $"PING round-trip took too long: {rtt.TotalMilliseconds}ms");

        Assert.True(sup.Shutdown(timeoutMs: 5000),
            "kernel did not exit within 5s of SHUTDOWN");
        Assert.True(sup.Process.HasExited);
        Assert.Equal(0, sup.Process.ExitCode);
    }

    [Fact]
    public void Multiple_Pings_All_Succeed()
    {
        if (ShouldSkipForPlatform()) return;

        var python = DiscoverPython();
        var pythonPath = DiscoverEmbeddedPath();
        Assert.False(string.IsNullOrEmpty(python), "python not found");
        Assert.False(string.IsNullOrEmpty(pythonPath), "embedded path not found");

        using var sup = KernelSupervisor.StartPython(python!, pythonPath!);

        // Health-check cadence the supervisor will emit in production. Each
        // PING is independent (own nonce), so a missed nonce echo would surface
        // as an InvalidOperationException, not a silent pass.
        for (var i = 0; i < 10; i++)
        {
            var rtt = sup.Ping(timeoutMs: 1000);
            Assert.True(rtt < TimeSpan.FromSeconds(1));
        }

        Assert.True(sup.Shutdown(timeoutMs: 5000));
        Assert.Equal(0, sup.Process.ExitCode);
    }

    [Fact]
    public void Dispose_Without_Shutdown_Still_Reaps_Process()
    {
        if (ShouldSkipForPlatform()) return;

        var python = DiscoverPython();
        var pythonPath = DiscoverEmbeddedPath();
        Assert.False(string.IsNullOrEmpty(python), "python not found");
        Assert.False(string.IsNullOrEmpty(pythonPath), "embedded path not found");

        var sup = KernelSupervisor.StartPython(python!, pythonPath!);
        var pid = sup.Process.Id;
        Assert.False(sup.Process.HasExited);

        sup.Dispose();

        // Dispose disposes the underlying Process object too — the captured
        // reference is no longer queryable. Verify the OS process is gone by
        // looking it up by PID. This is the "no orphaned python.exe after
        // Excel closes" guarantee from the roadmap, exercised in the unhappy
        // path where the caller forgot to call Shutdown explicitly.
        Assert.False(IsProcessAlive(pid),
            $"python child (pid {pid}) still alive after KernelSupervisor.Dispose");
    }

    private static bool IsProcessAlive(int pid)
    {
        try
        {
            using var probe = System.Diagnostics.Process.GetProcessById(pid);
            return !probe.HasExited;
        }
        catch (ArgumentException)
        {
            // No process with that PID — it's been reaped, which is what we want.
            return false;
        }
    }

    // -------------------------------------------------------------------------
    // Discovery helpers
    // -------------------------------------------------------------------------

    private static string? DiscoverPython()
    {
        var candidates = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? new[] { "python.exe", "python3.exe" }
            : new[] { "python3", "python" };

        var pathEnv = Environment.GetEnvironmentVariable("PATH") ?? "";
        foreach (var dir in pathEnv.Split(Path.PathSeparator))
        {
            if (string.IsNullOrWhiteSpace(dir)) continue;
            foreach (var name in candidates)
            {
                var full = Path.Combine(dir, name);
                if (File.Exists(full)) return full;
            }
        }
        return null;
    }

    private static string? DiscoverEmbeddedPath()
    {
        // Walk up from the test binary location until we find the repo's
        // embedded/ directory containing the kernel package. The test
        // assembly lives somewhere under tests/.../bin/...; the repo root is
        // 4-6 levels up depending on configuration.
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        for (var i = 0; i < 8 && dir != null; i++)
        {
            var candidate = Path.Combine(dir.FullName, "embedded", "pyexcel", "kernel", "__main__.py");
            if (File.Exists(candidate))
                return Path.Combine(dir.FullName, "embedded");
            dir = dir.Parent;
        }
        return null;
    }

    /// <summary>
    /// Returns true if the test should bail (pass silently) on this platform.
    /// xUnit 2.x has no in-test dynamic skip; the pragmatic alternative is
    /// an early return that reports as a pass. The Windows CI lane still
    /// runs every non-integration test in this assembly via the same
    /// dotnet-test invocation.
    /// </summary>
    private static bool ShouldSkipForPlatform()
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows);
    }
}
