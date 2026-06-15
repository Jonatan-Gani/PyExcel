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
        Assert.Equal(OrientationDecision.Ask, r.Decision);
        Assert.True(r.Ask);
        Assert.False(r.IsInvalid);
    }

    [Theory]
    [InlineData(1, 5)]   // a wide row
    [InlineData(1, 2)]
    [InlineData(1, 100)]
    public void Resolve_SingleRow_Horizontal(int rows, int cols)
    {
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.Equal(OrientationDecision.Resolved, r.Decision);
        Assert.Equal(ListOrientation.Horizontal, r.Orientation);
    }

    [Theory]
    [InlineData(5, 1)]   // a tall column
    [InlineData(2, 1)]
    [InlineData(100, 1)]
    public void Resolve_SingleColumn_Vertical(int rows, int cols)
    {
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.Equal(OrientationDecision.Resolved, r.Decision);
        Assert.Equal(ListOrientation.Vertical, r.Orientation);
    }

    [Theory]
    [InlineData(2, 2)]   // a square block
    [InlineData(3, 3)]
    [InlineData(2, 4)]   // wider block
    [InlineData(4, 2)]   // taller block
    public void Resolve_TwoDimensionalBlock_IsInvalid(int rows, int cols)
    {
        // A 1-D list can't fill a 2-D block unambiguously — the caller must reject it.
        var r = OrientationResolver.Resolve(rows, cols);
        Assert.Equal(OrientationDecision.Invalid, r.Decision);
        Assert.True(r.IsInvalid);
        Assert.False(r.Ask);
    }
}
