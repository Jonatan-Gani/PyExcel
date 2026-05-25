using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Tests for <see cref="ScriptDirectoryWatcher"/>. <see cref="FileSystemWatcher"/>
/// events fire on a thread-pool worker, so all assertions wait for the
/// callback to land rather than reading state synchronously.
/// </summary>
public class ScriptDirectoryWatcherTests : IDisposable
{
    private readonly string _dir;
    private readonly ConcurrentQueue<IReadOnlyList<string>> _snapshots = new();
    private readonly ManualResetEventSlim _signal = new(initialState: false);

    public ScriptDirectoryWatcherTests()
    {
        _dir = Path.Combine(
            Path.GetTempPath(),
            "pyexcel-scripts-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        _signal.Dispose();
        try { Directory.Delete(_dir, recursive: true); } catch { /* best-effort */ }
    }

    private ScriptDirectoryWatcher StartWatcher()
    {
        return new ScriptDirectoryWatcher(_dir, snapshot =>
        {
            _snapshots.Enqueue(snapshot);
            _signal.Set();
        });
    }

    /// <summary>Wait for at least one new snapshot to arrive and return it.
    /// Times out fast so test failures don't sit forever.</summary>
    private IReadOnlyList<string> WaitForSnapshot(int timeoutMs = 2000)
    {
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        while (DateTime.UtcNow < deadline)
        {
            if (_snapshots.TryDequeue(out var s)) return s;
            _signal.Wait(50);
            _signal.Reset();
        }
        throw new TimeoutException($"no snapshot within {timeoutMs}ms");
    }

    // -------------------------------------------------------------------------

    [Fact]
    public void InitialSnapshot_PushedSynchronously()
    {
        File.WriteAllText(Path.Combine(_dir, "alpha.py"), "");
        File.WriteAllText(Path.Combine(_dir, "beta.py"), "");

        using var w = StartWatcher();

        // Initial snapshot lands inside the constructor — already in the
        // queue by the time StartWatcher returns.
        var first = WaitForSnapshot();
        Assert.Equal(new[] { "alpha", "beta" }, first);
    }

    [Fact]
    public void AddingFile_TriggersSnapshot()
    {
        using var w = StartWatcher();
        _ = WaitForSnapshot();  // drain initial empty snapshot

        File.WriteAllText(Path.Combine(_dir, "new.py"), "");

        // Take snapshots until we see "new" in the list (the watcher may
        // fire multiple events for a single create on some platforms).
        var deadline = DateTime.UtcNow.AddSeconds(3);
        while (DateTime.UtcNow < deadline)
        {
            var s = WaitForSnapshot();
            if (s.Count == 1 && s[0] == "new") return;
        }
        Assert.Fail("snapshot containing 'new' never arrived");
    }

    [Fact]
    public void DeletingFile_TriggersSnapshot()
    {
        var path = Path.Combine(_dir, "doomed.py");
        File.WriteAllText(path, "");

        using var w = StartWatcher();
        _ = WaitForSnapshot();  // initial = {"doomed"}

        File.Delete(path);

        var deadline = DateTime.UtcNow.AddSeconds(3);
        while (DateTime.UtcNow < deadline)
        {
            var s = WaitForSnapshot();
            if (s.Count == 0) return;
        }
        Assert.Fail("snapshot reflecting the deletion never arrived");
    }

    [Fact]
    public void Refresh_PushesCurrentList()
    {
        File.WriteAllText(Path.Combine(_dir, "x.py"), "");
        using var w = StartWatcher();
        _ = WaitForSnapshot();  // initial = {"x"}

        // No file change — Refresh should still produce a snapshot.
        w.Refresh();
        var snap = WaitForSnapshot();
        Assert.Equal(new[] { "x" }, snap);
    }

    [Fact]
    public void NonPyFile_IsIgnored()
    {
        using var w = StartWatcher();
        _ = WaitForSnapshot();  // initial = {}

        File.WriteAllText(Path.Combine(_dir, "readme.txt"), "");

        // Refresh forces a poll so we see the current state without
        // relying on FileSystemWatcher to filter the .txt event for us.
        w.Refresh();
        var snap = WaitForSnapshot();
        Assert.Empty(snap);
    }

    [Fact]
    public void Constructor_NonexistentDirectory_Throws()
    {
        Assert.Throws<DirectoryNotFoundException>(() =>
            new ScriptDirectoryWatcher("/no/such/path", _ => { }));
    }

    [Fact]
    public void Dispose_IsIdempotent()
    {
        var w = StartWatcher();
        w.Dispose();
        w.Dispose();  // second call must not throw
    }
}
