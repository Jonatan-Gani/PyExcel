using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class PastePreflightTests
{
    // -------------------------------------------------------------------------
    // Footprint — dimensions of the decoded payload
    // -------------------------------------------------------------------------

    [Fact]
    public void Footprint_Null_IsZero()
    {
        var (rows, cols) = PastePreflight.Footprint(null);
        Assert.Equal(0, rows);
        Assert.Equal(0, cols);
    }

    [Fact]
    public void Footprint_TwoDimensionalTable_UsesItsOwnDimensions()
    {
        var table = new object?[3, 4];
        var (rows, cols) = PastePreflight.Footprint(table);
        Assert.Equal(3, rows);
        Assert.Equal(4, cols);
    }

    [Fact]
    public void Footprint_OneDimensionalVector_IsSingleRow()
    {
        var vec = new object?[5];
        var (rows, cols) = PastePreflight.Footprint(vec);
        Assert.Equal(1, rows);
        Assert.Equal(5, cols);
    }

    [Fact]
    public void Footprint_Scalar_IsOneByOne()
    {
        var (rows, cols) = PastePreflight.Footprint(42);
        Assert.Equal(1, rows);
        Assert.Equal(1, cols);
    }

    [Fact]
    public void Footprint_ScalarString_IsOneByOne()
    {
        var (rows, cols) = PastePreflight.Footprint("hello");
        Assert.Equal(1, rows);
        Assert.Equal(1, cols);
    }

    [Fact]
    public void Footprint_EmptyTable_IsZero()
    {
        // 0×0 happens when a scripts returns a DataFrame that decoded to
        // an empty table — the service should short-circuit in that case.
        var empty = new object?[0, 0];
        var (rows, cols) = PastePreflight.Footprint(empty);
        Assert.Equal(0, rows);
        Assert.Equal(0, cols);
    }

    [Fact]
    public void Footprint_EmptyVector_IsZeroWide()
    {
        var (rows, cols) = PastePreflight.Footprint(new object?[0]);
        Assert.Equal(1, rows);
        Assert.Equal(0, cols);
    }

    // -------------------------------------------------------------------------
    // RangeHasContent — does the target hold anything the paste would clobber?
    // -------------------------------------------------------------------------

    [Fact]
    public void RangeHasContent_Null_IsFalse()
    {
        Assert.False(PastePreflight.RangeHasContent(null));
    }

    [Fact]
    public void RangeHasContent_ScalarNull_IsFalse()
    {
        // The COM hand-back for a single empty cell.
        object? value = null;
        Assert.False(PastePreflight.RangeHasContent(value));
    }

    [Fact]
    public void RangeHasContent_ScalarEmptyString_IsFalse()
    {
        // Excel sometimes returns "" for a cell that displays as empty
        // but has been explicitly cleared with `Range.Value2 = ""`.
        Assert.False(PastePreflight.RangeHasContent(""));
    }

    [Fact]
    public void RangeHasContent_ScalarNumber_IsTrue()
    {
        Assert.True(PastePreflight.RangeHasContent(42.0));
    }

    [Fact]
    public void RangeHasContent_ScalarZero_IsTrue()
    {
        // 0 is still a value — overwriting it is destructive.
        Assert.True(PastePreflight.RangeHasContent(0.0));
    }

    [Fact]
    public void RangeHasContent_ScalarNonEmptyString_IsTrue()
    {
        Assert.True(PastePreflight.RangeHasContent("foo"));
    }

    [Fact]
    public void RangeHasContent_ScalarBool_IsTrue()
    {
        Assert.True(PastePreflight.RangeHasContent(true));
        Assert.True(PastePreflight.RangeHasContent(false));
    }

    [Fact]
    public void RangeHasContent_TableAllNull_IsFalse()
    {
        var table = new object?[3, 4];
        Assert.False(PastePreflight.RangeHasContent(table));
    }

    [Fact]
    public void RangeHasContent_TableAllEmptyStrings_IsFalse()
    {
        var table = new object?[2, 2];
        for (int i = 0; i < 2; i++)
            for (int j = 0; j < 2; j++)
                table[i, j] = "";
        Assert.False(PastePreflight.RangeHasContent(table));
    }

    [Fact]
    public void RangeHasContent_TableOneNonEmptyCell_IsTrue()
    {
        var table = new object?[3, 4];
        table[1, 2] = "hello";
        Assert.True(PastePreflight.RangeHasContent(table));
    }

    [Fact]
    public void RangeHasContent_TableOneZero_IsTrue()
    {
        var table = new object?[2, 2];
        table[0, 0] = 0.0;
        Assert.True(PastePreflight.RangeHasContent(table));
    }

    [Fact]
    public void RangeHasContent_TableMixedNullAndStrings_IsTrue()
    {
        var table = new object?[2, 3];
        table[1, 2] = "x";
        Assert.True(PastePreflight.RangeHasContent(table));
    }

    [Fact]
    public void RangeHasContent_VectorAllNull_IsFalse()
    {
        var vec = new object?[5];
        Assert.False(PastePreflight.RangeHasContent(vec));
    }

    [Fact]
    public void RangeHasContent_VectorWithContent_IsTrue()
    {
        var vec = new object?[3];
        vec[1] = 1.5;
        Assert.True(PastePreflight.RangeHasContent(vec));
    }

    [Fact]
    public void RangeHasContent_VectorAllEmptyStrings_IsFalse()
    {
        var vec = new object?[3] { "", "", "" };
        Assert.False(PastePreflight.RangeHasContent(vec));
    }

    [Fact]
    public void RangeHasContent_OneBasedComArray_StillCorrect()
    {
        // Excel hands back 1-based arrays. Build one explicitly and pin
        // that PastePreflight honours the lower bounds via GetLowerBound.
        var oneBased = System.Array.CreateInstance(
            typeof(object),
            new[] { 2, 3 },
            new[] { 1, 1 });
        // Set [1,1]..[2,3] all null first.
        oneBased.SetValue("hit", 2, 2);
        Assert.True(PastePreflight.RangeHasContent(oneBased));

        var emptyOneBased = System.Array.CreateInstance(
            typeof(object),
            new[] { 2, 2 },
            new[] { 1, 1 });
        Assert.False(PastePreflight.RangeHasContent(emptyOneBased));
    }
}
