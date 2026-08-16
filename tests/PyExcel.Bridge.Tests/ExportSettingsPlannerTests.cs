using System;
using System.IO;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ExportSettingsPlannerTests
{
    private static readonly DateTime Stamp = new(2026, 6, 16, 14, 30, 0);

    private static ExportSettings Settings(
        string? source = "A1:C10",
        string? folder = null,
        string? baseName = "report",
        ExportFileType fileType = ExportFileType.Csv,
        ExportTimestampStyle timestamp = ExportTimestampStyle.None)
        => new(source, folder, baseName, fileType, timestamp);

    // -------------------------------------------------------------------------
    // ComposeFileName
    // -------------------------------------------------------------------------

    [Fact]
    public void ComposeFileName_PlainCsv()
        => Assert.Equal("report.csv", ExportSettingsPlanner.ComposeFileName(Settings(), Stamp));

    [Fact]
    public void ComposeFileName_Tsv_UsesTsvExtension()
        => Assert.Equal("report.tsv",
            ExportSettingsPlanner.ComposeFileName(Settings(fileType: ExportFileType.Tsv), Stamp));

    [Theory]
    [InlineData(ExportTimestampStyle.DateAndTime, "report_2026-06-16_14-30-00.csv")]
    [InlineData(ExportTimestampStyle.DateOnly, "report_2026-06-16.csv")]
    [InlineData(ExportTimestampStyle.Compact, "report_20260616-143000.csv")]
    public void ComposeFileName_Stamped_AppendsStyleFormat(ExportTimestampStyle style, string expected)
        => Assert.Equal(expected,
            ExportSettingsPlanner.ComposeFileName(Settings(timestamp: style), Stamp));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ComposeFileName_BlankBaseName_FallsBackToDefault(string? baseName)
        => Assert.Equal("export.csv",
            ExportSettingsPlanner.ComposeFileName(Settings(baseName: baseName), Stamp));

    [Fact]
    public void ComposeFileName_TypedExtension_NotDoubled()
    {
        // The user typed "report.csv" as the base name — we must not produce
        // "report.csv.csv".
        Assert.Equal("report.csv",
            ExportSettingsPlanner.ComposeFileName(Settings(baseName: "report.csv"), Stamp));
        Assert.Equal("report.tsv",
            ExportSettingsPlanner.ComposeFileName(
                Settings(baseName: "report.TSV", fileType: ExportFileType.Tsv), Stamp));
    }

    // -------------------------------------------------------------------------
    // SanitizeBaseName
    // -------------------------------------------------------------------------

    [Fact]
    public void SanitizeBaseName_StripsReservedCharacters()
        => Assert.Equal("ab_cd",
            ExportSettingsPlanner.SanitizeBaseName("a<b>_c:d"));  // < > : dropped

    [Fact]
    public void SanitizeBaseName_StripsPathSeparators()
        => Assert.Equal("abc", ExportSettingsPlanner.SanitizeBaseName("a/b\\c"));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void SanitizeBaseName_Blank_IsEmpty(string? raw)
        => Assert.Equal(string.Empty, ExportSettingsPlanner.SanitizeBaseName(raw));

    // -------------------------------------------------------------------------
    // PreviewPattern
    // -------------------------------------------------------------------------

    [Fact]
    public void PreviewPattern_NoStamp_IsJustNameAndExtension()
        => Assert.Equal("report.csv", ExportSettingsPlanner.PreviewPattern(Settings()));

    [Fact]
    public void PreviewPattern_Stamped_HasStablePlaceholder()
        => Assert.Equal("report_{timestamp}.csv",
            ExportSettingsPlanner.PreviewPattern(Settings(timestamp: ExportTimestampStyle.DateAndTime)));

    [Fact]
    public void PreviewPattern_BlankName_UsesDefault()
        => Assert.Equal("export.csv", ExportSettingsPlanner.PreviewPattern(Settings(baseName: null)));

    // -------------------------------------------------------------------------
    // Resolve
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Resolve_BlankSource_Throws(string? source)
    {
        var ex = Assert.Throws<FormatException>(
            () => ExportSettingsPlanner.Resolve(Settings(source: source), Stamp, "/wb"));
        Assert.Contains("source range", ex.Message);
    }

    [Fact]
    public void Resolve_BlankFolder_SavesIntoWorkbookDirectory()
    {
        var basis = Environment.OSVersion.Platform == PlatformID.Win32NT ? @"C:\wb" : "/wb";
        var plan = ExportSettingsPlanner.Resolve(
            Settings(folder: null, timestamp: ExportTimestampStyle.None), Stamp, basis);

        Assert.Equal("A1:C10", plan.SourceRangeAddress);
        Assert.Equal(',', plan.Delimiter);
        Assert.Equal(Path.GetFullPath(Path.Combine(basis, "report.csv")), plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Resolve_RelativeFolder_ResolvesAgainstWorkbookDirectory()
    {
        var basis = Environment.OSVersion.Platform == PlatformID.Win32NT ? @"C:\wb" : "/wb";
        var plan = ExportSettingsPlanner.Resolve(Settings(folder: "exports"), Stamp, basis);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(basis, "exports", "report.csv")),
            plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Resolve_AbsoluteFolder_Wins()
    {
        var folder = Environment.OSVersion.Platform == PlatformID.Win32NT ? @"C:\out" : "/out";
        var plan = ExportSettingsPlanner.Resolve(Settings(folder: folder), Stamp, "/wb");
        Assert.Equal(Path.GetFullPath(Path.Combine(folder, "report.csv")), plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Resolve_Tsv_PicksTabDelimiterAndExtension()
    {
        var plan = ExportSettingsPlanner.Resolve(
            Settings(folder: null, fileType: ExportFileType.Tsv), Stamp, "/wb");
        Assert.Equal('\t', plan.Delimiter);
        Assert.EndsWith("report.tsv", plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Resolve_Stamped_ProducesUniqueNameInPath()
    {
        var plan = ExportSettingsPlanner.Resolve(
            Settings(folder: null, timestamp: ExportTimestampStyle.DateAndTime), Stamp, "/wb");
        Assert.EndsWith("report_2026-06-16_14-30-00.csv", plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Resolve_TrimsSourceRange()
    {
        var plan = ExportSettingsPlanner.Resolve(Settings(source: "  Sheet1!A1:B2  "), Stamp, "/wb");
        Assert.Equal("Sheet1!A1:B2", plan.SourceRangeAddress);
    }
}
