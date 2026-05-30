using System;
using System.Collections.Generic;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Unit tests for <see cref="ErrorService"/> and
/// <see cref="KernelErrorRecord.FormatForClipboard"/>.
/// </summary>
public class ErrorServiceTests
{
    // -------------------------------------------------------------------------
    // ErrorService — per-workbook + global slot, ErrorChanged event
    // -------------------------------------------------------------------------

    [Fact]
    public void GetLast_EmptyService_ReturnsNull()
    {
        var svc = new ErrorService();
        Assert.Null(svc.GetLast("anywhere"));
        Assert.Null(svc.GetLast(null));
    }

    [Fact]
    public void Record_PerWorkbook_IsScopedToThatKey()
    {
        var svc = new ErrorService();
        var record = MakeRecord("PY.RUN", "Exception", "ValueError");

        svc.Record("wbA", record);

        Assert.Same(record, svc.GetLast("wbA"));
        Assert.Null(svc.GetLast("wbB"));  // other workbook unaffected
    }

    [Fact]
    public void Record_GlobalSlot_FallsBackWhenWorkbookHasNoOwnError()
    {
        var svc = new ErrorService();
        var globalRecord = MakeRecord("Kernel boot", "HostError", "InvalidOperationException");

        svc.Record(workbookKey: null, globalRecord);

        // Workbook with no per-workbook entry sees the global one as a fallback.
        Assert.Same(globalRecord, svc.GetLast("wbA"));
        Assert.Same(globalRecord, svc.GetLast(null));
    }

    [Fact]
    public void Record_PerWorkbook_OverridesGlobalFallback()
    {
        var svc = new ErrorService();
        var globalRecord = MakeRecord("Kernel boot", "HostError", "InvalidOperationException");
        var wbRecord = MakeRecord("PY.RUN", "Exception", "ValueError");

        svc.Record(null, globalRecord);
        svc.Record("wbA", wbRecord);

        // Per-workbook wins for wbA …
        Assert.Same(wbRecord, svc.GetLast("wbA"));
        // … and the global slot still backstops every other workbook.
        Assert.Same(globalRecord, svc.GetLast("wbB"));
    }

    [Fact]
    public void Record_TwicePerWorkbook_ReplacesPriorRecord()
    {
        var svc = new ErrorService();
        var first = MakeRecord("PY.RUN", "Exception", "ValueError");
        var second = MakeRecord("PY.RUN", "Exception", "KeyError");

        svc.Record("wbA", first);
        svc.Record("wbA", second);

        Assert.Same(second, svc.GetLast("wbA"));
    }

    [Fact]
    public void Clear_PerWorkbook_RemovesOnlyThatKey()
    {
        var svc = new ErrorService();
        var record = MakeRecord("PY.RUN", "Exception", "ValueError");
        svc.Record("wbA", record);
        svc.Record("wbB", record);

        svc.Clear("wbA");

        Assert.Null(svc.GetLast("wbA"));
        Assert.Same(record, svc.GetLast("wbB"));
    }

    [Fact]
    public void Clear_GlobalSlot_LeavesPerWorkbookEntriesAlone()
    {
        var svc = new ErrorService();
        var wbRecord = MakeRecord("PY.RUN", "Exception", "ValueError");
        var globalRecord = MakeRecord("Kernel boot", "HostError", "OOM");
        svc.Record("wbA", wbRecord);
        svc.Record(null, globalRecord);

        svc.Clear(null);

        Assert.Same(wbRecord, svc.GetLast("wbA"));
        Assert.Null(svc.GetLast("wbB"));  // no per-workbook entry, global gone
    }

    [Fact]
    public void Record_FiresErrorChangedWithMatchingKey()
    {
        var svc = new ErrorService();
        var fired = new List<string?>();
        svc.ErrorChanged += (_, e) => fired.Add(e.WorkbookKey);

        svc.Record("wbA", MakeRecord("PY.RUN", "Exception", "ValueError"));
        svc.Record(null, MakeRecord("Kernel boot", "HostError", "OOM"));

        Assert.Equal(new string?[] { "wbA", null }, fired);
    }

    [Fact]
    public void Clear_FiresErrorChangedEvenWhenSlotWasEmpty()
    {
        // The ribbon's repaint hook is the only signal that "the buttons
        // should disable now" — firing on a no-op Clear means callers
        // never have to bookend their Clear call with a manual check.
        var svc = new ErrorService();
        var fired = new List<string?>();
        svc.ErrorChanged += (_, e) => fired.Add(e.WorkbookKey);

        svc.Clear("never-recorded");

        Assert.Single(fired);
        Assert.Equal("never-recorded", fired[0]);
    }

    [Fact]
    public void Record_NullRecord_Throws()
    {
        var svc = new ErrorService();
        Assert.Throws<ArgumentNullException>(() => svc.Record("wbA", null!));
    }

    // -------------------------------------------------------------------------
    // KernelErrorRecord.FormatForClipboard — stable layout users will paste
    // -------------------------------------------------------------------------

    [Fact]
    public void FormatForClipboard_IncludesEveryField()
    {
        var record = new KernelErrorRecord(
            Timestamp: new DateTimeOffset(2026, 5, 30, 14, 23, 5, TimeSpan.Zero),
            Source: "PY.RUN",
            Code: "Exception",
            PythonType: "ValueError",
            Message: "bad shape",
            PythonTraceback: "Traceback (most recent call last):\n  File \"x.py\", line 1\nValueError: bad shape",
            ScriptPath: "C:\\scripts\\x.py");

        var text = record.FormatForClipboard();

        Assert.Contains("[2026-05-30 14:23:05] PY.RUN", text);
        Assert.Contains("code:    Exception", text);
        Assert.Contains("type:    ValueError", text);
        Assert.Contains("message: bad shape", text);
        Assert.Contains("script:  C:\\scripts\\x.py", text);
        Assert.Contains("traceback:", text);
        // Traceback lines are indented so the block visually attaches
        // to the header.
        Assert.Contains("    Traceback (most recent call last):", text);
        Assert.Contains("    ValueError: bad shape", text);
    }

    [Fact]
    public void FormatForClipboard_OmitsRedundantTypeWhenSameAsCode()
    {
        // Host-side faults have Code == PythonType — no point printing
        // both. The format suppresses the duplicate.
        var record = new KernelErrorRecord(
            Timestamp: DateTimeOffset.MinValue,
            Source: "PY.RUN",
            Code: "HostError",
            PythonType: "HostError",
            Message: "boom",
            PythonTraceback: "",
            ScriptPath: null);

        var text = record.FormatForClipboard();

        Assert.Contains("code:    HostError", text);
        Assert.DoesNotContain("type:    HostError", text);
    }

    [Fact]
    public void FormatForClipboard_OmitsEmptyOptionalFields()
    {
        var record = new KernelErrorRecord(
            Timestamp: DateTimeOffset.MinValue,
            Source: "PY.RUN",
            Code: "Exception",
            PythonType: "ValueError",
            Message: "x",
            PythonTraceback: "",
            ScriptPath: null);

        var text = record.FormatForClipboard();

        Assert.DoesNotContain("script:", text);
        Assert.DoesNotContain("traceback:", text);
    }

    private static KernelErrorRecord MakeRecord(string source, string code, string pyType)
        => new(
            Timestamp: DateTimeOffset.UtcNow,
            Source: source,
            Code: code,
            PythonType: pyType,
            Message: "test message",
            PythonTraceback: "",
            ScriptPath: null);
}
