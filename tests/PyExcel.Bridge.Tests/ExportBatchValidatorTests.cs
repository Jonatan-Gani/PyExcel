using System;
using System.Collections.Generic;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ExportBatchValidatorTests
{
    private static List<ExportJob> Jobs(params (string src, string tgt)[] rows)
    {
        var list = new List<ExportJob>();
        foreach (var (src, tgt) in rows) list.Add(new ExportJob(src, tgt));
        return list;
    }

    [Fact]
    public void Validate_NullJobs_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => ExportBatchValidator.Validate(null!, null));
    }

    [Fact]
    public void Validate_NoRows_Fails()
    {
        var r = ExportBatchValidator.Validate(Jobs(), null);
        Assert.False(r.IsValid);
        Assert.Contains("at least one", r.ErrorMessage!);
    }

    [Fact]
    public void Validate_ValidRows_OkTrimmed()
    {
        var r = ExportBatchValidator.Validate(
            Jobs(("  A1:C10 ", " out1.csv "), ("D1", "out2.tsv")), null);
        Assert.True(r.IsValid);
        Assert.Equal(2, r.Jobs.Count);
        Assert.Equal("A1:C10", r.Jobs[0].SourceRange);
        Assert.Equal("out1.csv", r.Jobs[0].TargetPath);
    }

    [Fact]
    public void Validate_BlankSource_FailsWithRowNumber()
    {
        var r = ExportBatchValidator.Validate(
            Jobs(("A1", "out1.csv"), ("", "out2.csv")), null);
        Assert.False(r.IsValid);
        Assert.Contains("Row 2", r.ErrorMessage!);
    }

    [Fact]
    public void Validate_ExcelTarget_FailsWithRowNumber()
    {
        var r = ExportBatchValidator.Validate(
            Jobs(("A1", "out.xlsx")), null);
        Assert.False(r.IsValid);
        Assert.Contains("Row 1", r.ErrorMessage!);
    }

    [Fact]
    public void Validate_DuplicateTarget_Fails()
    {
        var r = ExportBatchValidator.Validate(
            Jobs(("A1", "out.csv"), ("B1", "OUT.csv")), null);
        Assert.False(r.IsValid);
        Assert.Contains("same file", r.ErrorMessage!);
    }
}
