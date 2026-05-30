using System;
using System.IO;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ImportPlannerTests
{
    // -------------------------------------------------------------------------
    // Field validation
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankImportInput_Throws(string? input)
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create(input, "A1", workbookDirectory: null));
        Assert.Contains("Input", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankImportOutput_Throws(string? output)
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create("data.csv", output, workbookDirectory: null));
        Assert.Contains("Output", ex.Message);
    }

    [Fact]
    public void Create_TrimsLeadingAndTrailingWhitespace()
    {
        var plan = ImportPlanner.Create(
            "  data.csv  ",
            "  Sheet1!A1  ",
            workbookDirectory: "/wb");
        Assert.Equal("Sheet1!A1", plan.TargetRangeAddress);
        // Whitespace stripped before the path was joined to /wb.
        Assert.EndsWith("data.csv", plan.AbsoluteSourcePath);
    }

    // -------------------------------------------------------------------------
    // Path resolution
    // -------------------------------------------------------------------------

    [Fact]
    public void ResolvePath_AbsoluteIsPreserved()
    {
        var abs = OperatingSystem_IsWindows() ? @"C:\data\foo.csv" : "/data/foo.csv";
        var resolved = ImportPlanner.ResolvePath(abs, workbookDirectory: "/elsewhere");
        Assert.Equal(Path.GetFullPath(abs), resolved);
    }

    [Fact]
    public void ResolvePath_RelativeJoinsToWorkbookDir()
    {
        var basis = OperatingSystem_IsWindows() ? @"C:\wb" : "/wb";
        var resolved = ImportPlanner.ResolvePath("data.csv", workbookDirectory: basis);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(basis, "data.csv")),
            resolved);
    }

    [Fact]
    public void ResolvePath_RelativeNoWorkbookDir_UsesCurrentDirectory()
    {
        var resolved = ImportPlanner.ResolvePath("data.csv", workbookDirectory: null);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "data.csv")),
            resolved);
    }

    [Fact]
    public void ResolvePath_BlankSource_Throws()
    {
        Assert.Throws<ArgumentException>(() => ImportPlanner.ResolvePath("", "/wb"));
    }

    // -------------------------------------------------------------------------
    // Delimiter detection
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("foo.csv", ',')]
    [InlineData("foo.txt", ',')]
    [InlineData("foo", ',')]
    [InlineData("foo.CSV", ',')]
    public void DetectDelimiter_DefaultsToComma(string path, char expected)
    {
        Assert.Equal(expected, ImportPlanner.DetectDelimiter(path));
    }

    [Theory]
    [InlineData("foo.tsv")]
    [InlineData("foo.TSV")]
    [InlineData("/path/to/foo.tsv")]
    public void DetectDelimiter_TsvExtension_ReturnsTab(string path)
    {
        Assert.Equal('\t', ImportPlanner.DetectDelimiter(path));
    }

    [Theory]
    [InlineData("foo.xlsx")]
    [InlineData("foo.xls")]
    [InlineData("foo.xlsm")]
    [InlineData("foo.xlsb")]
    [InlineData("foo.ods")]
    [InlineData("foo.XLSX")]
    public void DetectDelimiter_BinaryFormats_Throws(string path)
    {
        var ex = Assert.Throws<FormatException>(() => ImportPlanner.DetectDelimiter(path));
        Assert.Contains("not yet supported", ex.Message);
    }

    [Fact]
    public void DetectDelimiter_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ImportPlanner.DetectDelimiter(null!));
    }

    // -------------------------------------------------------------------------
    // End-to-end Create() composition
    // -------------------------------------------------------------------------

    [Fact]
    public void Create_TsvExtension_PicksTabDelimiter()
    {
        var plan = ImportPlanner.Create("data.tsv", "A1", workbookDirectory: "/wb");
        Assert.Equal('\t', plan.Delimiter);
    }

    [Fact]
    public void Create_CsvExtension_PicksCommaDelimiter()
    {
        var plan = ImportPlanner.Create("data.csv", "A1", workbookDirectory: "/wb");
        Assert.Equal(',', plan.Delimiter);
    }

    [Fact]
    public void Create_XlsxExtension_Throws()
    {
        // The Excel-format import is a separate Phase 5 follow-up; the
        // planner has to reject it explicitly so the user gets a clean
        // error instead of a CSV mis-parse.
        Assert.Throws<FormatException>(
            () => ImportPlanner.Create("data.xlsx", "A1", workbookDirectory: "/wb"));
    }

    /// <summary>Tiny helper — `OperatingSystem.IsWindows()` lives in
    /// net5+; the test project targets net8 so it's available, but
    /// rolling our own keeps the test source independent of which TFM
    /// the test assembly happens to be on.</summary>
    private static bool OperatingSystem_IsWindows()
        => Environment.OSVersion.Platform is PlatformID.Win32NT
            or PlatformID.Win32S
            or PlatformID.Win32Windows
            or PlatformID.WinCE;
}
