using System;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="ChartTraceData"/> shapes a trace's raw JSON cells into
/// COM-ready series arrays — the cross-platform port of v1
/// <c>chartBuilder.bas</c>'s <c>BuildXYForCategories</c> /
/// <c>BuildXYNumericFlexible</c> rules (category detection, ISO-date →
/// OADate conversion, null/NaN pair dropping, histogram label
/// preference).
/// </summary>
public class ChartTraceDataTests
{
    private static ChartTraceSpec Trace(
        object?[] x, object?[] y,
        string seriesType = "scatter",
        string[]? text = null)
        => new(
            Id: 1,
            X: x,
            Y: y,
            Text: text,
            Size: null,
            Style: new ChartTraceStyle(
                SeriesType: seriesType,
                Name: "t",
                SecondaryAxis: false,
                Visible: true,
                Line: new ChartLineStyle("#000000", "solid", null),
                Marker: new ChartMarkerStyle(6, "#000000", "circle"),
                FillColor: null,
                FillOpacity: null));

    // -------------------------------------------------------------------------
    // IsCategorical
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("bar", true)]
    [InlineData("column", true)]
    [InlineData("pie", true)]
    [InlineData("histogram", true)]
    [InlineData("scatter", false)]
    [InlineData("scatter_lines", false)]
    [InlineData("line", false)]
    [InlineData("area", false)]
    [InlineData("bubble", false)]
    public void IsCategorical_ClassifiesSeriesTypes(string seriesType, bool expected)
    {
        Assert.Equal(expected, ChartTraceData.IsCategorical(seriesType));
    }

    // -------------------------------------------------------------------------
    // Numeric series
    // -------------------------------------------------------------------------

    [Fact]
    public void Shape_NumericXY_PassesThrough()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1L, 2.5 }, new object?[] { 10L, 20L }))!;
        Assert.Equal(new object[] { 1.0, 2.5 }, shaped.XValues);
        Assert.Equal(new object[] { 10.0, 20.0 }, shaped.YValues);
        Assert.False(shaped.XIsDate);
    }

    [Fact]
    public void Shape_NullPairs_Dropped()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1L, null, 3L }, new object?[] { 10L, 20L, null }))!;
        Assert.Equal(new object[] { 1.0 }, shaped.XValues);
        Assert.Equal(new object[] { 10.0 }, shaped.YValues);
    }

    [Fact]
    public void Shape_IsoDateX_ConvertsToOADateAndFlags()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { "2026-01-02", "2026-01-03" },
            new object?[] { 1L, 2L }))!;
        Assert.True(shaped.XIsDate);
        Assert.Equal(new DateTime(2026, 1, 2).ToOADate(), (double)shaped.XValues![0]);
        Assert.Equal(new DateTime(2026, 1, 3).ToOADate(), (double)shaped.XValues![1]);
    }

    [Fact]
    public void Shape_IsoDateTimeX_Parses()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { "2026-01-02T06:00:00" }, new object?[] { 1L }))!;
        Assert.True(shaped.XIsDate);
        Assert.Equal(
            new DateTime(2026, 1, 2, 6, 0, 0).ToOADate(),
            (double)shaped.XValues![0]);
    }

    [Fact]
    public void Shape_OneNonDateStringX_FlipsWholeAxisToCategories()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1L, "Q2", 3L }, new object?[] { 10L, 20L, 30L }))!;
        Assert.Equal(new object[] { "1", "Q2", "3" }, shaped.XValues);
        Assert.False(shaped.XIsDate);
    }

    [Fact]
    public void Shape_MixedDatesAndNumbers_StaysNumeric()
    {
        // A numeric x mixed with ISO dates is still a numeric axis (dates
        // become serials); only a non-date string flips to categories.
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { "2026-01-02", 45000L }, new object?[] { 1L, 2L }))!;
        Assert.Equal(new DateTime(2026, 1, 2).ToOADate(), (double)shaped.XValues![0]);
        Assert.Equal(45000.0, (double)shaped.XValues![1]);
    }

    [Fact]
    public void Shape_EmptyX_YieldsNullXValues()
    {
        var shaped = ChartTraceData.Shape(Trace(
            Array.Empty<object?>(), new object?[] { 1L, 2L }))!;
        Assert.Null(shaped.XValues);
        Assert.Equal(new object[] { 1.0, 2.0 }, shaped.YValues);
    }

    [Fact]
    public void Shape_NoUsableData_ReturnsNull()
    {
        Assert.Null(ChartTraceData.Shape(Trace(
            new object?[] { 1L }, new object?[] { (object?)null })));
        Assert.Null(ChartTraceData.Shape(Trace(
            Array.Empty<object?>(), Array.Empty<object?>())));
    }

    [Fact]
    public void Shape_NaNY_Dropped()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1L, 2L }, new object?[] { double.NaN, 5.0 }))!;
        Assert.Equal(new object[] { 2.0 }, shaped.XValues);
        Assert.Equal(new object[] { 5.0 }, shaped.YValues);
    }

    // -------------------------------------------------------------------------
    // Categorical series
    // -------------------------------------------------------------------------

    [Fact]
    public void Shape_CategoricalColumn_XBecomesStrings()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { "a", "b" }, new object?[] { 1L, 2L },
            seriesType: "column"))!;
        Assert.Equal(new object[] { "a", "b" }, shaped.XValues);
        Assert.Equal(new object[] { 1.0, 2.0 }, shaped.YValues);
        Assert.False(shaped.XIsDate);
    }

    [Fact]
    public void Shape_CategoricalNumericX_Stringified()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1L, 2.5 }, new object?[] { 10L, 20L },
            seriesType: "bar"))!;
        Assert.Equal(new object[] { "1", "2.5" }, shaped.XValues);
    }

    [Fact]
    public void Shape_CategoricalNonNumericY_PairDropped()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { "a", "b", "c" },
            new object?[] { 1L, "oops", 3L },
            seriesType: "pie"))!;
        Assert.Equal(new object[] { "a", "c" }, shaped.XValues);
        Assert.Equal(new object[] { 1.0, 3.0 }, shaped.YValues);
    }

    [Fact]
    public void Shape_CategoricalAllYInvalid_ReturnsNull()
    {
        Assert.Null(ChartTraceData.Shape(Trace(
            new object?[] { "a" }, new object?[] { "not a number" },
            seriesType: "column")));
    }

    // -------------------------------------------------------------------------
    // Histogram label preference
    // -------------------------------------------------------------------------

    [Fact]
    public void Shape_HistogramWithMatchingText_UsesLabelsAsCategories()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1.5, 4.5 },          // bin midpoints
            new object?[] { 3L, 7L },            // counts
            seriesType: "histogram",
            text: new[] { "0:3", "3:6" }))!;     // kernel's lo:hi labels
        Assert.Equal(new object[] { "0:3", "3:6" }, shaped.XValues);
        Assert.Equal(new object[] { 3.0, 7.0 }, shaped.YValues);
    }

    [Fact]
    public void Shape_HistogramWithMismatchedText_FallsBackToMidpoints()
    {
        var shaped = ChartTraceData.Shape(Trace(
            new object?[] { 1.5, 4.5 },
            new object?[] { 3L, 7L },
            seriesType: "histogram",
            text: new[] { "only one label" }))!;
        Assert.Equal(new object[] { "1.5", "4.5" }, shaped.XValues);
    }

    [Fact]
    public void Shape_NullTrace_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ChartTraceData.Shape(null!));
    }
}
