using PyExcel.Forms;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class RangeAddressValidatorTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_Blank_Fails(string? input)
    {
        var r = RangeAddressValidator.Validate(input);
        Assert.False(r.IsValid);
        Assert.Null(r.Address);
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("A1:C10")]
    [InlineData("Sheet1!A1:C10")]
    public void Validate_SingleRange_Ok(string input)
    {
        var r = RangeAddressValidator.Validate(input);
        Assert.True(r.IsValid);
        Assert.Equal(input, r.Address);
    }

    [Fact]
    public void Validate_Trims()
    {
        var r = RangeAddressValidator.Validate("   A1:B2   ");
        Assert.True(r.IsValid);
        Assert.Equal("A1:B2", r.Address);
    }

    [Fact]
    public void Validate_MultipleRanges_Fails()
    {
        var r = RangeAddressValidator.Validate("A1;B2");
        Assert.False(r.IsValid);
        Assert.Contains("single range", r.ErrorMessage!);
    }

    [Fact]
    public void Validate_NamePrefix_Fails()
    {
        var r = RangeAddressValidator.Validate("prices=A1:C10");
        Assert.False(r.IsValid);
        Assert.Contains("name=", r.ErrorMessage!);
    }
}
