using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class OrientationResolverTests
{
    [Theory]
    [InlineData(1, 1)]
    [InlineData(0, 0)]
    [InlineData(1, 0)]
    [InlineData(0, 1)]
    public void Resolve_SingleOrDegenerate_AsksUser(int rows, int cols)
    {
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.True(r.Ask);
    }

    [Theory]
    [InlineData(1, 5)]   // a wide row
    [InlineData(3, 3)]   // square -> horizontal (matches v1 >=)
    [InlineData(2, 4)]
    public void Resolve_WiderOrSquare_Horizontal(int rows, int cols)
    {
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.False(r.Ask);
        Assert.Equal(ListOrientation.Horizontal, r.Orientation);
    }

    [Theory]
    [InlineData(5, 1)]   // a tall column
    [InlineData(4, 2)]
    public void Resolve_Taller_Vertical(int rows, int cols)
    {
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.False(r.Ask);
        Assert.Equal(ListOrientation.Vertical, r.Orientation);
    }
}
