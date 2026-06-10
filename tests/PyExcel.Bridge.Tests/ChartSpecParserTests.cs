using System;
using System.Linq;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// The C# half of the chart-spec contract: <see cref="ChartSpecParser"/>
/// must accept everything <c>pyexcel.kernel.chart.plotly_figure_to_spec</c>
/// emits and reject malformed documents with actionable messages. Pairs
/// with <c>tests/kernel/test_chart.py</c>, which pins the producing side.
/// </summary>
public class ChartSpecParserTests
{
    /// <summary>A representative document of everything the Python
    /// producer emits — used as the parse-everything smoke test and as
    /// a base for the focused cases below.</summary>
    private const string FullSpec = @"
        {
          ""version"": 1,
          ""chart_type"": ""xy"",
          ""title"": ""Quarterly"",
          ""x_axis"": {""title"": ""Time"", ""min"": 0, ""max"": 10, ""log_scale"": false},
          ""y_axis"": {""title"": ""Price"", ""min"": null, ""max"": null, ""log_scale"": true},
          ""legend"": {""visible"": true, ""position"": ""bottom""},
          ""barmode"": ""group"",
          ""traces"": [
            {
              ""id"": 1,
              ""x"": [1, 2.5, ""2026-01-02"", null],
              ""y"": [10, 20, 30, null],
              ""text"": [""a"", ""b"", ""c"", ""d""],
              ""size"": null,
              ""style"": {
                ""series_type"": ""scatter_lines_markers"",
                ""name"": ""prices"",
                ""axis_group"": ""secondary"",
                ""visible"": true,
                ""line"": {""color"": ""#ff0000"", ""dash"": ""dash"", ""width"": 2.5},
                ""marker"": {""size"": 8, ""color"": ""#00ff00"", ""shape"": ""square""},
                ""fill_color"": null,
                ""fill_opacity"": 0.5
              }
            }
          ],
          ""annotations"": [
            {
              ""id"": 1,
              ""type"": ""threshold"",
              ""axis"": ""y"",
              ""value"": 25,
              ""style"": {""color"": ""#0000ff"", ""width"": 2, ""dash"": ""dot"", ""opacity"": 0.8},
              ""label"": ""limit""
            }
          ]
        }";

    // -------------------------------------------------------------------------
    // Happy path
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_FullSpec_RoundTripsEveryField()
    {
        var doc = ChartSpecParser.Parse(FullSpec);

        Assert.Equal(1, doc.Version);
        Assert.Equal("xy", doc.ChartType);
        Assert.Equal("Quarterly", doc.Title);

        Assert.Equal("Time", doc.XAxis.Title);
        Assert.Equal(0, doc.XAxis.Min);
        Assert.Equal(10, doc.XAxis.Max);
        Assert.False(doc.XAxis.LogScale);

        Assert.Equal("Price", doc.YAxis.Title);
        Assert.Null(doc.YAxis.Min);
        Assert.Null(doc.YAxis.Max);
        Assert.True(doc.YAxis.LogScale);

        Assert.True(doc.Legend.Visible);
        Assert.Equal("bottom", doc.Legend.Position);
        Assert.Equal("group", doc.BarMode);

        var trace = Assert.Single(doc.Traces);
        Assert.Equal(1, trace.Id);
        Assert.Equal(new object?[] { 1L, 2.5, "2026-01-02", null }, trace.X);
        Assert.Equal(new object?[] { 10L, 20L, 30L, null }, trace.Y);
        Assert.Equal(new[] { "a", "b", "c", "d" }, trace.Text);
        Assert.Null(trace.Size);

        Assert.Equal("scatter_lines_markers", trace.Style.SeriesType);
        Assert.Equal("prices", trace.Style.Name);
        Assert.True(trace.Style.SecondaryAxis);
        Assert.True(trace.Style.Visible);
        Assert.Equal(new ChartLineStyle("#ff0000", "dash", 2.5), trace.Style.Line);
        Assert.Equal(new ChartMarkerStyle(8, "#00ff00", "square"), trace.Style.Marker);
        Assert.Null(trace.Style.FillColor);
        Assert.Equal(0.5, trace.Style.FillOpacity);

        var ann = Assert.Single(doc.Annotations);
        Assert.Equal("threshold", ann.Type);
        Assert.Equal("y", ann.Axis);
        Assert.Equal(25L, ann.Value);
        Assert.Equal(new ChartAnnotationStyle("#0000ff", 2, "dot", 0.8), ann.Style);
        Assert.Equal("limit", ann.Label);
    }

    [Fact]
    public void Parse_MinimalSpec_AppliesDefaults()
    {
        var doc = ChartSpecParser.Parse(@"{""version"": 1}");

        Assert.Equal("xy", doc.ChartType);
        Assert.Equal(string.Empty, doc.Title);
        Assert.Equal(new ChartAxisSpec("", null, null, false), doc.XAxis);
        Assert.Equal(new ChartAxisSpec("", null, null, false), doc.YAxis);
        Assert.True(doc.Legend.Visible);
        Assert.Equal("right", doc.Legend.Position);
        Assert.Equal("group", doc.BarMode);
        Assert.Empty(doc.Traces);
        Assert.Empty(doc.Annotations);
    }

    [Theory]
    [InlineData("xy")]
    [InlineData("line")]
    [InlineData("bar")]
    [InlineData("area")]
    [InlineData("pie")]
    public void Parse_EveryDocumentedChartType_Accepted(string chartType)
    {
        var doc = ChartSpecParser.Parse(
            @"{""version"": 1, ""chart_type"": """ + chartType + @"""}");
        Assert.Equal(chartType, doc.ChartType);
    }

    [Theory]
    [InlineData("scatter")]
    [InlineData("scatter_lines")]
    [InlineData("scatter_lines_markers")]
    [InlineData("line")]
    [InlineData("column")]
    [InlineData("bar")]
    [InlineData("area")]
    [InlineData("pie")]
    [InlineData("histogram")]
    [InlineData("bubble")]
    public void Parse_EveryDocumentedSeriesType_Accepted(string seriesType)
    {
        var doc = ChartSpecParser.Parse(TraceWithSeriesType(seriesType));
        Assert.Equal(seriesType, doc.Traces[0].Style.SeriesType);
    }

    [Fact]
    public void Parse_TraceStyleDefaults_Applied()
    {
        var doc = ChartSpecParser.Parse(@"
            {""version"": 1, ""traces"": [
              {""id"": 3, ""y"": [1], ""style"": {""series_type"": ""scatter""}}
            ]}");
        var style = doc.Traces[0].Style;
        Assert.Equal("Series 3", style.Name);
        Assert.False(style.SecondaryAxis);
        Assert.True(style.Visible);
        Assert.Equal(new ChartLineStyle("#000000", "solid", null), style.Line);
        Assert.Equal(new ChartMarkerStyle(6.0, "#000000", "circle"), style.Marker);
        Assert.Null(style.FillColor);
        Assert.Null(style.FillOpacity);
    }

    [Fact]
    public void Parse_BubbleSizes_ParsedAsNumbers()
    {
        var doc = ChartSpecParser.Parse(@"
            {""version"": 1, ""traces"": [
              {""id"": 1, ""x"": [1, 2], ""y"": [3, 4], ""size"": [10, 20.5],
               ""style"": {""series_type"": ""bubble""}}
            ]}");
        Assert.Equal(new[] { 10.0, 20.5 }, doc.Traces[0].Size);
    }

    // -------------------------------------------------------------------------
    // Rejections — every message must name the offending element
    // -------------------------------------------------------------------------

    [Fact]
    public void Parse_NullInput_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ChartSpecParser.Parse(null!));
    }

    [Fact]
    public void Parse_InvalidJson_ThrowsWithContext()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse("{not json"));
        Assert.Contains("not valid JSON", ex.Message);
    }

    [Fact]
    public void Parse_NonObjectRoot_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse("[1, 2]"));
        Assert.Contains("root must be a JSON object", ex.Message);
    }

    [Fact]
    public void Parse_MissingVersion_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => ChartSpecParser.Parse(@"{""chart_type"": ""xy""}"));
        Assert.Contains("version", ex.Message);
    }

    [Fact]
    public void Parse_UnsupportedVersion_NamesBothVersions()
    {
        var ex = Assert.Throws<FormatException>(
            () => ChartSpecParser.Parse(@"{""version"": 99}"));
        Assert.Contains("99", ex.Message);
        Assert.Contains("version 1", ex.Message);
    }

    [Fact]
    public void Parse_UnsupportedChartType_NamesSupportedSet()
    {
        var ex = Assert.Throws<FormatException>(
            () => ChartSpecParser.Parse(@"{""version"": 1, ""chart_type"": ""radar""}"));
        Assert.Contains("radar", ex.Message);
        Assert.Contains("pie", ex.Message);
    }

    [Fact]
    public void Parse_UnsupportedSeriesType_NamesTraceAndSupportedSet()
    {
        var ex = Assert.Throws<FormatException>(
            () => ChartSpecParser.Parse(TraceWithSeriesType("heatmap")));
        Assert.Contains("heatmap", ex.Message);
        Assert.Contains("trace 1", ex.Message);
        Assert.Contains("histogram", ex.Message);
    }

    [Fact]
    public void Parse_TraceMissingStyle_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""traces"": [{""id"": 1, ""y"": [1]}]}"));
        Assert.Contains("style", ex.Message);
    }

    [Fact]
    public void Parse_TraceMissingSeriesType_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""traces"": [{""id"": 1, ""y"": [1], ""style"": {}}]}"));
        Assert.Contains("series_type", ex.Message);
    }

    [Fact]
    public void Parse_TraceMissingId_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""traces"": [{""y"": [1], ""style"": {""series_type"": ""line""}}]}"));
        Assert.Contains("id", ex.Message);
    }

    [Fact]
    public void Parse_DuplicateTraceIds_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(@"
            {""version"": 1, ""traces"": [
              {""id"": 7, ""y"": [1], ""style"": {""series_type"": ""line""}},
              {""id"": 7, ""y"": [2], ""style"": {""series_type"": ""line""}}
            ]}"));
        Assert.Contains("duplicate trace id 7", ex.Message);
    }

    [Fact]
    public void Parse_MismatchedXYLengths_NamesCounts()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(@"
            {""version"": 1, ""traces"": [
              {""id"": 1, ""x"": [1, 2, 3], ""y"": [1], ""style"": {""series_type"": ""line""}}
            ]}"));
        Assert.Contains("3", ex.Message);
        Assert.Contains("1", ex.Message);
    }

    [Fact]
    public void Parse_TracesNotAnArray_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""traces"": ""nope""}"));
        Assert.Contains("array", ex.Message);
    }

    [Fact]
    public void Parse_NonNumericAxisBound_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""x_axis"": {""min"": ""zero""}}"));
        Assert.Contains("min", ex.Message);
        Assert.Contains("number", ex.Message);
    }

    [Fact]
    public void Parse_BadLegendPosition_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(
            @"{""version"": 1, ""legend"": {""position"": ""top""}}"));
        Assert.Contains("top", ex.Message);
    }

    [Fact]
    public void Parse_UnsupportedAnnotationType_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(@"
            {""version"": 1, ""annotations"": [
              {""id"": 1, ""type"": ""arrow"", ""axis"": ""x"", ""value"": 1}
            ]}"));
        Assert.Contains("arrow", ex.Message);
        Assert.Contains("event_line", ex.Message);
    }

    [Fact]
    public void Parse_AnnotationMissingValue_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(@"
            {""version"": 1, ""annotations"": [
              {""id"": 1, ""type"": ""threshold"", ""axis"": ""y""}
            ]}"));
        Assert.Contains("value", ex.Message);
    }

    [Fact]
    public void Parse_AnnotationBadAxis_Throws()
    {
        var ex = Assert.Throws<FormatException>(() => ChartSpecParser.Parse(@"
            {""version"": 1, ""annotations"": [
              {""id"": 1, ""type"": ""threshold"", ""axis"": ""z"", ""value"": 1}
            ]}"));
        Assert.Contains("'z'", ex.Message);
    }

    [Fact]
    public void SupportedSeriesTypes_MatchPythonSide()
    {
        // Pinned copy of pyexcel.kernel.chart.SUPPORTED_SERIES_TYPES —
        // if one side gains a type the other must too.
        Assert.Equal(
            new[]
            {
                "scatter", "scatter_lines", "scatter_lines_markers", "line",
                "column", "bar", "area", "pie", "histogram", "bubble",
            },
            ChartSpecParser.SupportedSeriesTypes.ToArray());
    }

    private static string TraceWithSeriesType(string seriesType) => @"
        {""version"": 1, ""traces"": [
          {""id"": 1, ""x"": [1], ""y"": [2], ""style"": {""series_type"": """ + seriesType + @"""}}
        ]}";
}
