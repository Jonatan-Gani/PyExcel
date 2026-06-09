using System;
using System.IO;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ExportPlannerTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankExportInput_Throws(string? input)
    {
        var ex = Assert.Throws<FormatException>(
            () => ExportPlanner.Create(input, "out.csv", workbookDirectory: null));
        Assert.Contains("Input", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankExportOutput_Throws(string? output)
    {
        var ex = Assert.Throws<FormatException>(
            () => ExportPlanner.Create("A1:C10", output, workbookDirectory: null));
        Assert.Contains("Output", ex.Message);
    }

    [Fact]
    public void Create_HappyPath_PreservesSourceAddressAndPicksCommaDelimiter()
    {
        var plan = ExportPlanner.Create("Sheet1!A1:C10", "out.csv", workbookDirectory: "/wb");
        Assert.Equal("Sheet1!A1:C10", plan.SourceRangeAddress);
        Assert.Equal(',', plan.Delimiter);
        Assert.EndsWith("out.csv", plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Create_TsvExtension_PicksTabDelimiter()
    {
        var plan = ExportPlanner.Create("A1:C10", "out.tsv", workbookDirectory: "/wb");
        Assert.Equal('\t', plan.Delimiter);
    }

    [Fact]
    public void Create_RelativeTargetResolvedAgainstWorkbookDir()
    {
        var basis = Environment.OSVersion.Platform == PlatformID.Win32NT
            ? @"C:\wb"
            : "/wb";
        var plan = ExportPlanner.Create("A1", "out.csv", workbookDirectory: basis);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(basis, "out.csv")),
            plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Create_TrimsLeadingAndTrailingWhitespace()
    {
        var plan = ExportPlanner.Create(
            "  A1:C10  ",
            "  out.csv  ",
            workbookDirectory: "/wb");
        Assert.Equal("A1:C10", plan.SourceRangeAddress);
        Assert.EndsWith("out.csv", plan.AbsoluteTargetPath);
    }

    [Fact]
    public void Create_XlsxExtension_Throws()
    {
        Assert.Throws<FormatException>(
            () => ExportPlanner.Create("A1", "out.xlsx", workbookDirectory: "/wb"));
    }
}
