using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Unit tests for <see cref="RunArchive"/>: directory layout, retention
/// cap, manifest round-trip, concurrent-write safety.
/// </summary>
public class RunArchiveTests : IDisposable
{
    private readonly string _root;

    public RunArchiveTests()
    {
        _root = Path.Combine(Path.GetTempPath(), "PyExcelRunArchiveTests_" + Guid.NewGuid().ToString("N"));
    }

    public void Dispose()
    {
        try { if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true); }
        catch { /* best-effort cleanup */ }
    }

    // -------------------------------------------------------------------------
    // Construction
    // -------------------------------------------------------------------------

    [Fact]
    public void Ctor_NullOrEmptyRoot_Throws()
    {
        Assert.Throws<ArgumentException>(() => new RunArchive(""));
        Assert.Throws<ArgumentException>(() => new RunArchive("   "));
        Assert.Throws<ArgumentException>(() => new RunArchive(null!));
    }

    [Fact]
    public void Ctor_NegativeMaxRuns_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new RunArchive(_root, maxRuns: -1));
    }

    [Fact]
    public void Ctor_DoesNotCreateRootDirectory()
    {
        // Lazy creation — constructing the service shouldn't poke the FS.
        // The default service is built at app-domain start; a fast-path
        // exit before any run shouldn't leave an empty PyExcel/runs dir
        // littered on the user's machine.
        _ = new RunArchive(_root);
        Assert.False(Directory.Exists(_root));
    }

    // -------------------------------------------------------------------------
    // Archive — directory layout
    // -------------------------------------------------------------------------

    [Fact]
    public void Archive_Success_WritesInputsOutputAndManifest()
    {
        var archive = new RunArchive(_root);
        var entry = SuccessEntry(
            timestamp: new DateTimeOffset(2026, 5, 30, 14, 0, 0, 123, TimeSpan.Zero),
            inputs: new[] { new byte[] { 1, 2, 3 }, new byte[] { 9, 8 } },
            output: new byte[] { 4, 5, 6, 7 });

        var dir = archive.Archive(entry);

        Assert.True(Directory.Exists(dir));
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(Path.Combine(dir, "input_0.arrow")));
        Assert.Equal(new byte[] { 9, 8 }, File.ReadAllBytes(Path.Combine(dir, "input_1.arrow")));
        Assert.Equal(new byte[] { 4, 5, 6, 7 }, File.ReadAllBytes(Path.Combine(dir, "output.arrow")));
        Assert.True(File.Exists(Path.Combine(dir, "manifest.txt")));
        Assert.False(File.Exists(Path.Combine(dir, "error.txt")));
    }

    [Fact]
    public void Archive_RunIdEmbedsTimestamp()
    {
        var archive = new RunArchive(_root);
        var entry = SuccessEntry(
            timestamp: new DateTimeOffset(2026, 5, 30, 14, 0, 0, 123, TimeSpan.Zero));

        var dir = archive.Archive(entry);

        // Directory leaf starts with the UTC timestamp in lexicographic
        // form so List()/Prune() can sort by name without re-reading any
        // manifests.
        Assert.StartsWith("20260530T140000123_", Path.GetFileName(dir));
    }

    [Fact]
    public void Archive_NoOutput_OmitsOutputFile()
    {
        // A user function returning None gives RunResult.IsEmpty=true; the
        // caller passes output=null. Archive should record the run with
        // no output.arrow file.
        var archive = new RunArchive(_root);
        var entry = SuccessEntry(output: null);

        var dir = archive.Archive(entry);

        Assert.False(File.Exists(Path.Combine(dir, "output.arrow")));
        Assert.True(File.Exists(Path.Combine(dir, "manifest.txt")));
    }

    [Fact]
    public void Archive_Error_WritesErrorTxtFromFormatForClipboard()
    {
        var archive = new RunArchive(_root);
        var error = new KernelErrorRecord(
            Timestamp: new DateTimeOffset(2026, 5, 30, 14, 0, 0, 0, TimeSpan.Zero),
            Source: "PY.RUN",
            Code: "Exception",
            PythonType: "ValueError",
            Message: "bad shape",
            PythonTraceback: "Traceback (most recent call last):\n  File \"x.py\", line 1\nValueError: bad shape",
            ScriptPath: "C:\\scripts\\x.py");
        var entry = ErrorEntry(error);

        var dir = archive.Archive(entry);

        var errorTxt = File.ReadAllText(Path.Combine(dir, "error.txt"));
        Assert.Equal(error.FormatForClipboard(), errorTxt);
        Assert.False(File.Exists(Path.Combine(dir, "output.arrow")));
    }

    [Fact]
    public void Archive_NullInputBuffer_Throws()
    {
        var archive = new RunArchive(_root);
        var entry = new RunArchiveEntry(
            Timestamp: DateTimeOffset.UtcNow,
            WorkbookKey: null,
            ScriptPath: "x.py",
            Function: "transform",
            Source: "PY.RUN",
            Duration: TimeSpan.Zero,
            Inputs: new byte[][] { null! },
            Output: null,
            Error: null,
            Status: RunArchiveStatus.Success);

        Assert.Throws<ArgumentException>(() => archive.Archive(entry));
    }

    [Fact]
    public void Archive_NullEntry_Throws()
    {
        var archive = new RunArchive(_root);
        Assert.Throws<ArgumentNullException>(() => archive.Archive(null!));
    }

    // -------------------------------------------------------------------------
    // Manifest — header line by line
    // -------------------------------------------------------------------------

    [Fact]
    public void Manifest_IncludesEveryHeadlineField()
    {
        var archive = new RunArchive(_root);
        var entry = SuccessEntry(
            timestamp: new DateTimeOffset(2026, 5, 30, 14, 0, 0, 123, TimeSpan.Zero),
            workbookKey: "C:\\book.xlsx",
            scriptPath: "C:\\scripts\\transform.py",
            function: "compute",
            source: "Run Python button",
            duration: TimeSpan.FromMilliseconds(456),
            inputs: new[] { new byte[] { 1 }, new byte[] { 2 } });

        var dir = archive.Archive(entry);
        var manifest = File.ReadAllText(Path.Combine(dir, "manifest.txt"));

        Assert.Contains("RunId: 20260530T140000123_", manifest);
        Assert.Contains("TimestampUtc: 2026-05-30T14:00:00.1230000Z", manifest);
        Assert.Contains("DurationMs: 456", manifest);
        Assert.Contains("Source: Run Python button", manifest);
        Assert.Contains("WorkbookKey: C:\\book.xlsx", manifest);
        Assert.Contains("ScriptPath: C:\\scripts\\transform.py", manifest);
        Assert.Contains("Function: compute", manifest);
        Assert.Contains("InputCount: 2", manifest);
        Assert.Contains("Status: Success", manifest);
    }

    [Fact]
    public void Manifest_NullWorkbookKey_OmitsLine()
    {
        var archive = new RunArchive(_root);
        var entry = SuccessEntry(workbookKey: null);

        var dir = archive.Archive(entry);
        var manifest = File.ReadAllText(Path.Combine(dir, "manifest.txt"));

        // The line for WorkbookKey is suppressed entirely when there's no
        // workbook to record — parsers can detect "unbound run" by its
        // absence rather than by a sentinel value.
        Assert.DoesNotContain("WorkbookKey:", manifest);
    }

    [Fact]
    public void Manifest_ErrorRun_IncludesErrorCodeTypeMessage()
    {
        var archive = new RunArchive(_root);
        var error = new KernelErrorRecord(
            Timestamp: DateTimeOffset.UtcNow,
            Source: "PY.RUN",
            Code: "Exception",
            PythonType: "ValueError",
            Message: "bad shape",
            PythonTraceback: "tb",
            ScriptPath: "x.py");

        var dir = archive.Archive(ErrorEntry(error));
        var manifest = File.ReadAllText(Path.Combine(dir, "manifest.txt"));

        Assert.Contains("Status: Error", manifest);
        Assert.Contains("ErrorCode: Exception", manifest);
        Assert.Contains("ErrorType: ValueError", manifest);
        Assert.Contains("ErrorMessage: bad shape", manifest);
    }

    [Fact]
    public void Manifest_ErrorWithRedundantType_OmitsTypeLine()
    {
        // Host-side errors have Code == PythonType; the manifest mirrors
        // the FormatForClipboard suppression so a parser sees the same
        // information shape as the user's clipboard text.
        var archive = new RunArchive(_root);
        var error = new KernelErrorRecord(
            Timestamp: DateTimeOffset.UtcNow,
            Source: "PY.RUN",
            Code: "HostError",
            PythonType: "HostError",
            Message: "boom",
            PythonTraceback: "",
            ScriptPath: null);

        var dir = archive.Archive(ErrorEntry(error));
        var manifest = File.ReadAllText(Path.Combine(dir, "manifest.txt"));

        Assert.Contains("ErrorCode: HostError", manifest);
        Assert.DoesNotContain("ErrorType:", manifest);
    }

    [Fact]
    public void Manifest_MultilineErrorMessage_CollapsesToSingleLine()
    {
        // A multi-line error message would otherwise break the
        // "Key: Value per line" invariant of manifest.txt. The traceback
        // (which is intentionally multi-line) lives in error.txt instead.
        var archive = new RunArchive(_root);
        var error = new KernelErrorRecord(
            Timestamp: DateTimeOffset.UtcNow,
            Source: "PY.RUN",
            Code: "Exception",
            PythonType: "ValueError",
            Message: "line one\nline two\r\nline three",
            PythonTraceback: "",
            ScriptPath: null);

        var dir = archive.Archive(ErrorEntry(error));
        var manifestLines = File.ReadAllLines(Path.Combine(dir, "manifest.txt"));

        // Exactly one line should start with "ErrorMessage:".
        var msgLines = manifestLines.Where(l => l.StartsWith("ErrorMessage:", StringComparison.Ordinal)).ToList();
        Assert.Single(msgLines);
        Assert.DoesNotContain('\n', msgLines[0]);
    }

    // -------------------------------------------------------------------------
    // Retention cap
    // -------------------------------------------------------------------------

    [Fact]
    public void Archive_BeyondMaxRuns_PrunesOldest()
    {
        var archive = new RunArchive(_root, maxRuns: 3);
        var t0 = new DateTimeOffset(2026, 5, 30, 14, 0, 0, TimeSpan.Zero);

        // Five runs at strictly increasing timestamps. Lexicographic
        // ordering of the RunId prefix matches chronological order, so
        // the oldest two (#0 and #1) should be evicted.
        for (var i = 0; i < 5; i++)
            archive.Archive(SuccessEntry(timestamp: t0.AddSeconds(i)));

        var dirs = Directory.GetDirectories(_root)
            .Select(Path.GetFileName)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToList();

        Assert.Equal(3, dirs.Count);
        // The oldest survivor's RunId starts with t0+2s = 14:00:02.
        Assert.StartsWith("20260530T140002000_", dirs[0]);
        Assert.StartsWith("20260530T140003000_", dirs[1]);
        Assert.StartsWith("20260530T140004000_", dirs[2]);
    }

    [Fact]
    public void Archive_MaxRunsZero_PrunesEverything()
    {
        // Edge case: a caller that explicitly wants archiving off can set
        // maxRuns=0 and the directory stays empty after each call. Useful
        // for tests that exercise the wiring but want no on-disk footprint.
        var archive = new RunArchive(_root, maxRuns: 0);
        archive.Archive(SuccessEntry());

        Assert.Empty(Directory.GetDirectories(_root));
    }

    [Fact]
    public void Prune_NoOpUntilOverCap()
    {
        var archive = new RunArchive(_root, maxRuns: 3);
        archive.Archive(SuccessEntry());
        archive.Archive(SuccessEntry());

        archive.Prune();

        Assert.Equal(2, Directory.GetDirectories(_root).Length);
    }

    // -------------------------------------------------------------------------
    // List
    // -------------------------------------------------------------------------

    [Fact]
    public void List_EmptyRoot_ReturnsEmpty()
    {
        var archive = new RunArchive(_root);
        Assert.Empty(archive.List());
    }

    [Fact]
    public void List_ReturnsRunsNewestFirst()
    {
        var archive = new RunArchive(_root);
        var t0 = new DateTimeOffset(2026, 5, 30, 14, 0, 0, TimeSpan.Zero);
        archive.Archive(SuccessEntry(timestamp: t0, scriptPath: "first.py"));
        archive.Archive(SuccessEntry(timestamp: t0.AddSeconds(1), scriptPath: "second.py"));
        archive.Archive(SuccessEntry(timestamp: t0.AddSeconds(2), scriptPath: "third.py"));

        var listed = archive.List();

        Assert.Equal(3, listed.Count);
        Assert.Equal("third.py", listed[0].ScriptPath);
        Assert.Equal("second.py", listed[1].ScriptPath);
        Assert.Equal("first.py", listed[2].ScriptPath);
    }

    [Fact]
    public void List_ParsesManifestFields()
    {
        var archive = new RunArchive(_root);
        var ts = new DateTimeOffset(2026, 5, 30, 14, 0, 0, 123, TimeSpan.Zero);
        archive.Archive(SuccessEntry(
            timestamp: ts,
            workbookKey: "wb-key",
            scriptPath: "transform.py",
            source: "Run Python button"));

        var run = archive.List().Single();

        Assert.Equal(ts, run.Timestamp);
        Assert.Equal("wb-key", run.WorkbookKey);
        Assert.Equal("transform.py", run.ScriptPath);
        Assert.Equal("Run Python button", run.Source);
        Assert.Equal(RunArchiveStatus.Success, run.Status);
    }

    [Fact]
    public void List_SkipsCorruptManifest()
    {
        var archive = new RunArchive(_root);
        archive.Archive(SuccessEntry(scriptPath: "good.py"));

        // A directory that looks like a run but whose manifest is junk
        // should not crash List() — older schema, partial write, hand-
        // edited by the user. We silently skip it.
        var brokenDir = Path.Combine(_root, "20990101T000000000_deadbeef");
        Directory.CreateDirectory(brokenDir);
        File.WriteAllText(Path.Combine(brokenDir, "manifest.txt"), "this is not a valid manifest");

        var listed = archive.List();

        Assert.Single(listed);
        Assert.Equal("good.py", listed[0].ScriptPath);
    }

    // -------------------------------------------------------------------------
    // Concurrent writes — the archive must survive parallel runs without
    // corrupting its retention cap.
    // -------------------------------------------------------------------------

    [Fact]
    public async Task Archive_ConcurrentWrites_RespectMaxRunsCap()
    {
        const int maxRuns = 5;
        const int writers = 8;
        const int writesEach = 6;

        var archive = new RunArchive(_root, maxRuns: maxRuns);

        // Slight ts spread so directory names are distinct. The point of
        // this test is not deterministic ordering — it's that the on-disk
        // footprint stays within MaxRuns despite parallel Archive calls
        // hammering the prune step.
        var tasks = Enumerable.Range(0, writers).Select(w => Task.Run(() =>
        {
            for (var i = 0; i < writesEach; i++)
            {
                archive.Archive(SuccessEntry(
                    timestamp: DateTimeOffset.UtcNow.AddTicks(w * 1000 + i)));
                Thread.Sleep(1);
            }
        })).ToArray();

        await Task.WhenAll(tasks);

        var dirs = Directory.GetDirectories(_root);
        Assert.True(dirs.Length <= maxRuns,
            $"expected at most {maxRuns} archive directories, found {dirs.Length}");
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static RunArchiveEntry SuccessEntry(
        DateTimeOffset? timestamp = null,
        string? workbookKey = "C:\\book.xlsx",
        string scriptPath = "C:\\scripts\\transform.py",
        string function = "transform",
        string source = "PY.RUN",
        TimeSpan? duration = null,
        byte[][]? inputs = null,
        byte[]? output = null)
        => new(
            Timestamp: timestamp ?? new DateTimeOffset(2026, 5, 30, 14, 0, 0, TimeSpan.Zero),
            WorkbookKey: workbookKey,
            ScriptPath: scriptPath,
            Function: function,
            Source: source,
            Duration: duration ?? TimeSpan.FromMilliseconds(100),
            Inputs: inputs ?? new[] { new byte[] { 0x1 } },
            Output: output ?? new byte[] { 0x2 },
            Error: null,
            Status: RunArchiveStatus.Success);

    private static RunArchiveEntry ErrorEntry(KernelErrorRecord error)
        => new(
            Timestamp: new DateTimeOffset(2026, 5, 30, 14, 0, 0, TimeSpan.Zero),
            WorkbookKey: "C:\\book.xlsx",
            ScriptPath: "C:\\scripts\\transform.py",
            Function: "transform",
            Source: "PY.RUN",
            Duration: TimeSpan.FromMilliseconds(50),
            Inputs: new[] { new byte[] { 0x1 } },
            Output: null,
            Error: error,
            Status: RunArchiveStatus.Error);
}
