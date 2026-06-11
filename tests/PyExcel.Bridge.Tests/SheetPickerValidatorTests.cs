using System;
using PyExcel.Forms;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class SheetPickerValidatorTests
{
    private static readonly string[] Sheets = { "Inputs", "Q2 Data", "Summary" };

    [Fact]
    public void Validate_NullAvailableSheets_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => SheetPickerValidator.Validate("Inputs", null!));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_BlankSelection_Fails(string? selected)
    {
        var result = SheetPickerValidator.Validate(selected, Sheets);
        Assert.False(result.IsValid);
        Assert.Null(result.SelectedSheet);
        Assert.Contains("Select a sheet", result.ErrorMessage!);
    }

    [Fact]
    public void Validate_SelectionNotInWorkbook_Fails()
    {
        var result = SheetPickerValidator.Validate("Nope", Sheets);
        Assert.False(result.IsValid);
        Assert.Contains("not in this workbook", result.ErrorMessage!);
    }

    [Fact]
    public void Validate_ExactMatch_OkReturnsName()
    {
        var result = SheetPickerValidator.Validate("Q2 Data", Sheets);
        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
        Assert.Equal("Q2 Data", result.SelectedSheet);
    }

    [Fact]
    public void Validate_CaseInsensitiveMatch_ReturnsCanonicalCasing()
    {
        // Excel sheet names are unique case-insensitively; the picker must
        // hand back the workbook's own casing so the COM lookup matches.
        var result = SheetPickerValidator.Validate("summary", Sheets);
        Assert.True(result.IsValid);
        Assert.Equal("Summary", result.SelectedSheet);
    }

    [Fact]
    public void Validate_TrimsSelection()
    {
        var result = SheetPickerValidator.Validate("  Inputs  ", Sheets);
        Assert.True(result.IsValid);
        Assert.Equal("Inputs", result.SelectedSheet);
    }
}
