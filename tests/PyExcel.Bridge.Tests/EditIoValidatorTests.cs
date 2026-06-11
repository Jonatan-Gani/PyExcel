using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class EditIoValidatorTests
{
    // Pass a null workbook directory throughout — relative paths then
    // resolve against the process working directory, which is all the
    // planner needs for validation (no file I/O happens here).

    // -------------------------------------------------------------------------
    // Import
    // -------------------------------------------------------------------------

    [Fact]
    public void ValidateImport_ValidCsv_Ok()
    {
        var r = EditIoValidator.ValidateImport("data.csv", "A1", null);
        Assert.True(r.IsValid);
        Assert.Null(r.ErrorMessage);
        Assert.Equal("data.csv", r.Input);
        Assert.Equal("A1", r.Output);
    }

    [Fact]
    public void ValidateImport_ValidExcelWithSheet_Ok()
    {
        var r = EditIoValidator.ValidateImport("book.xlsx!Inputs", "Sheet1!B2", null);
        Assert.True(r.IsValid);
        Assert.Equal("book.xlsx!Inputs", r.Input);
        Assert.Equal("Sheet1!B2", r.Output);
    }

    [Fact]
    public void ValidateImport_TrimsValues()
    {
        var r = EditIoValidator.ValidateImport("  data.csv  ", "  A1  ", null);
        Assert.True(r.IsValid);
        Assert.Equal("data.csv", r.Input);
        Assert.Equal("A1", r.Output);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ValidateImport_BlankInput_Fails(string? input)
    {
        var r = EditIoValidator.ValidateImport(input, "A1", null);
        Assert.False(r.IsValid);
        Assert.Contains("Input", r.ErrorMessage!);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ValidateImport_BlankOutput_Fails(string? output)
    {
        var r = EditIoValidator.ValidateImport("data.csv", output, null);
        Assert.False(r.IsValid);
        Assert.Contains("Output", r.ErrorMessage!);
    }

    [Theory]
    [InlineData("legacy.xls")]
    [InlineData("sheet.ods")]
    public void ValidateImport_UnsupportedFormat_Fails(string input)
    {
        var r = EditIoValidator.ValidateImport(input, "A1", null);
        Assert.False(r.IsValid);
        Assert.NotNull(r.ErrorMessage);
    }

    // -------------------------------------------------------------------------
    // Export
    // -------------------------------------------------------------------------

    [Fact]
    public void ValidateExport_ValidCsvTarget_Ok()
    {
        var r = EditIoValidator.ValidateExport("A1:C10", "out.csv", null);
        Assert.True(r.IsValid);
        Assert.Equal("A1:C10", r.Input);
        Assert.Equal("out.csv", r.Output);
    }

    [Fact]
    public void ValidateExport_TrimsValues()
    {
        var r = EditIoValidator.ValidateExport("  A1:C10 ", " out.tsv ", null);
        Assert.True(r.IsValid);
        Assert.Equal("A1:C10", r.Input);
        Assert.Equal("out.tsv", r.Output);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ValidateExport_BlankInput_Fails(string? input)
    {
        var r = EditIoValidator.ValidateExport(input, "out.csv", null);
        Assert.False(r.IsValid);
        Assert.Contains("Input", r.ErrorMessage!);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ValidateExport_BlankOutput_Fails(string? output)
    {
        var r = EditIoValidator.ValidateExport("A1:C10", output, null);
        Assert.False(r.IsValid);
        Assert.Contains("Output", r.ErrorMessage!);
    }

    [Theory]
    [InlineData("book.xlsx")]
    [InlineData("book.xls")]
    [InlineData("sheet.ods")]
    public void ValidateExport_ExcelFormatTarget_Fails(string output)
    {
        var r = EditIoValidator.ValidateExport("A1:C10", output, null);
        Assert.False(r.IsValid);
        Assert.NotNull(r.ErrorMessage);
    }

    // -------------------------------------------------------------------------
    // Paste
    // -------------------------------------------------------------------------

    [Fact]
    public void ValidatePaste_ValidRange_Ok()
    {
        var r = EditIoValidator.ValidatePaste("A1");
        Assert.True(r.IsValid);
        Assert.Null(r.Input);
        Assert.Equal("A1", r.Output);
    }

    [Fact]
    public void ValidatePaste_TrimsRange()
    {
        var r = EditIoValidator.ValidatePaste("  Sheet1!D4  ");
        Assert.True(r.IsValid);
        Assert.Equal("Sheet1!D4", r.Output);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ValidatePaste_BlankRange_Fails(string? output)
    {
        var r = EditIoValidator.ValidatePaste(output);
        Assert.False(r.IsValid);
        Assert.Contains("target range", r.ErrorMessage!);
    }
}
