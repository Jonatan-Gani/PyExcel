using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.Excel;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class PastePlannerTests
{
    // -------------------------------------------------------------------------
    // Field validation
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankPasteOutput_Throws(string? output)
    {
        var ex = Assert.Throws<FormatException>(
            () => PastePlanner.Create(output, workbookKey: null, recentRuns: SomeRunWithOutput()));
        Assert.Contains("Output", ex.Message);
    }

    [Fact]
    public void Create_NullRunsList_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => PastePlanner.Create("A1", workbookKey: null, recentRuns: null!));
    }

    [Fact]
    public void Create_TrimsLeadingTrailingWhitespace()
    {
        var plan = PastePlanner.Create(
            "  Sheet1!A1  ",
            workbookKey: null,
            recentRuns: SomeRunWithOutput());
        Assert.Equal("Sheet1!A1", plan.TargetRangeAddress);
    }

    // -------------------------------------------------------------------------
    // Selection rule — newest output-bearing run wins
    // -------------------------------------------------------------------------

    [Fact]
    public void Create_NoRuns_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => PastePlanner.Create("A1", workbookKey: null, recentRuns: Array.Empty<ArchivedRun>()));
        Assert.Contains("no recent run", ex.Message);
    }

    [Fact]
    public void Create_NoRunsWithOutput_Throws()
    {
        // Two runs but neither produced a payload.
        var runs = new[]
        {
            Run("20260530T100000000_aaaa", hasOutput: false, workbookKey: null),
            Run("20260530T090000000_bbbb", hasOutput: false, workbookKey: null),
        };
        Assert.Throws<FormatException>(
            () => PastePlanner.Create("A1", workbookKey: null, recentRuns: runs));
    }

    [Fact]
    public void Create_NewestOutputBearingRunWins()
    {
        var runs = new[]
        {
            // Newest is HasOutput=false; the planner skips it and picks
            // the next-newest that has output.
            Run("20260530T120000000_cccc", hasOutput: false, workbookKey: null),
            Run("20260530T110000000_dddd", hasOutput: true,  workbookKey: null),
            Run("20260530T100000000_eeee", hasOutput: true,  workbookKey: null),
        };
        var plan = PastePlanner.Create("A1", workbookKey: null, recentRuns: runs);
        Assert.EndsWith("20260530T110000000_dddd" + Path.DirectorySeparatorChar + "output.arrow",
            plan.SourceArrowPath);
        Assert.Equal("20260530T110000000_dddd", plan.SourceRunId);
    }

    // -------------------------------------------------------------------------
    // Workbook-key preference
    // -------------------------------------------------------------------------

    [Fact]
    public void Create_PrefersSameWorkbookRunOverNewerUnboundRun()
    {
        var runs = new[]
        {
            // Newest is unbound but the workbook-keyed run is the user's
            // intent — pick the older one.
            Run("20260530T120000000_a", hasOutput: true, workbookKey: null),
            Run("20260530T110000000_b", hasOutput: true, workbookKey: "wb.xlsx"),
            Run("20260530T100000000_c", hasOutput: true, workbookKey: "other.xlsx"),
        };
        var plan = PastePlanner.Create("A1", workbookKey: "wb.xlsx", recentRuns: runs);
        Assert.Equal("20260530T110000000_b", plan.SourceRunId);
    }

    [Fact]
    public void Create_NoSameWorkbookRun_FallsBackToNewestOutputBearing()
    {
        // The user's workbook has no archived runs; pick the newest
        // output-bearing run from any source.
        var runs = new[]
        {
            Run("20260530T120000000_a", hasOutput: true, workbookKey: "other.xlsx"),
            Run("20260530T110000000_b", hasOutput: true, workbookKey: null),
        };
        var plan = PastePlanner.Create("A1", workbookKey: "fresh.xlsx", recentRuns: runs);
        Assert.Equal("20260530T120000000_a", plan.SourceRunId);
    }

    [Fact]
    public void Create_NullWorkbookKey_TakesNewestOutputBearing()
    {
        var runs = new[]
        {
            Run("20260530T120000000_a", hasOutput: true, workbookKey: "wb.xlsx"),
            Run("20260530T110000000_b", hasOutput: true, workbookKey: null),
        };
        var plan = PastePlanner.Create("A1", workbookKey: null, recentRuns: runs);
        // No preference applies — newest output-bearing wins outright.
        Assert.Equal("20260530T120000000_a", plan.SourceRunId);
    }

    // -------------------------------------------------------------------------
    // Plan composition
    // -------------------------------------------------------------------------

    [Fact]
    public void Create_SourcePathIsArchiveDirectoryOutputArrow()
    {
        var dir = Path.Combine(Path.GetTempPath(), "fake-archive", "20260530T100000000_xx");
        var run = new ArchivedRun(
            Directory: dir,
            RunId: "20260530T100000000_xx",
            Timestamp: DateTimeOffset.UtcNow,
            Status: RunArchiveStatus.Success,
            ScriptPath: "transform.py",
            WorkbookKey: null,
            Source: "PY.RUN",
            HasOutput: true);

        var plan = PastePlanner.Create("A1", workbookKey: null, recentRuns: new[] { run });
        Assert.Equal(Path.Combine(dir, "output.arrow"), plan.SourceArrowPath);
        Assert.Equal("A1", plan.TargetRangeAddress);
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static ArchivedRun Run(string runId, bool hasOutput, string? workbookKey)
        => new(
            Directory: Path.Combine(Path.GetTempPath(), "fake-archive", runId),
            RunId: runId,
            Timestamp: DateTimeOffset.UtcNow,
            Status: RunArchiveStatus.Success,
            ScriptPath: "transform.py",
            WorkbookKey: workbookKey,
            Source: "PY.RUN",
            HasOutput: hasOutput);

    private static IReadOnlyList<ArchivedRun> SomeRunWithOutput()
        => new[] { Run("20260530T100000000_aaaa", hasOutput: true, workbookKey: null) };
}
