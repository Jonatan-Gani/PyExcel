using System;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class SheetSelectionTests
{
    [Fact]
    public void Resolve_NullAvailable_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => SheetSelection.Resolve("Sheet1", null!));
    }

    [Fact]
    public void Resolve_PinnedSheet_AlwaysResolvedRegardlessOfCount()
    {
        var r = SheetSelection.Resolve("Q2", new[] { "A", "B", "C" });
        Assert.Equal(SheetResolutionKind.Resolved, r.Kind);
        Assert.Equal("Q2", r.Sheet);
    }

    [Fact]
    public void Resolve_PinnedSheet_Trimmed()
    {
        var r = SheetSelection.Resolve("  Q2  ", new[] { "A", "B" });
        Assert.Equal(SheetResolutionKind.Resolved, r.Kind);
        Assert.Equal("Q2", r.Sheet);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Resolve_NoPin_SingleSheet_ResolvesToThatSheet(string? pin)
    {
        var r = SheetSelection.Resolve(pin, new[] { "Only" });
        Assert.Equal(SheetResolutionKind.Resolved, r.Kind);
        Assert.Equal("Only", r.Sheet);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Resolve_NoPin_MultipleSheets_Prompts(string? pin)
    {
        var available = new[] { "A", "B", "C" };
        var r = SheetSelection.Resolve(pin, available);
        Assert.Equal(SheetResolutionKind.Prompt, r.Kind);
        Assert.Null(r.Sheet);
        Assert.Equal(available, r.AvailableSheets);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Resolve_NoPin_NoSheets_Empty(string? pin)
    {
        var r = SheetSelection.Resolve(pin, Array.Empty<string>());
        Assert.Equal(SheetResolutionKind.Empty, r.Kind);
        Assert.Null(r.Sheet);
        Assert.Empty(r.AvailableSheets);
    }

    [Fact]
    public void Resolve_PinnedSheet_EmptyWorkbook_StillResolvedPin()
    {
        // A pinned name wins even against an empty list — the COM lookup
        // then surfaces "not found", which is the right error for a pin
        // that doesn't exist.
        var r = SheetSelection.Resolve("Ghost", Array.Empty<string>());
        Assert.Equal(SheetResolutionKind.Resolved, r.Kind);
        Assert.Equal("Ghost", r.Sheet);
    }
}
