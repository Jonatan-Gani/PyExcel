using System.Collections.Generic;

namespace PyExcel.Excel;

/// <summary>
/// Typed model of the chart-spec JSON document (schema version 1) the
/// kernel emits for a Plotly figure. Produced by
/// <see cref="ChartSpecParser.Parse"/>; consumed by <c>ChartBuilder</c>.
/// The producing side and the schema reference live in
/// <c>embedded/pyexcel/kernel/chart.py</c>.
/// </summary>
public sealed record ChartSpecDocument(
    int Version,
    string ChartType,
    string Title,
    ChartAxisSpec XAxis,
    ChartAxisSpec YAxis,
    ChartLegendSpec Legend,
    string BarMode,
    IReadOnlyList<ChartTraceSpec> Traces,
    IReadOnlyList<ChartAnnotationSpec> Annotations);

/// <summary>One axis: optional title, optional explicit bounds, log flag.
/// Null bounds mean "let Excel auto-scale".</summary>
public sealed record ChartAxisSpec(
    string Title,
    double? Min,
    double? Max,
    bool LogScale);

/// <summary>Legend visibility + position ("right" or "bottom").</summary>
public sealed record ChartLegendSpec(
    bool Visible,
    string Position);

/// <summary>One data series. <see cref="X"/> cells are numbers, strings
/// (categories or ISO-8601 dates), or null; <see cref="Y"/> cells are
/// numbers or null. <see cref="Text"/> carries per-point labels;
/// <see cref="Size"/> carries bubble sizes. Both are null when absent.</summary>
public sealed record ChartTraceSpec(
    int Id,
    IReadOnlyList<object?> X,
    IReadOnlyList<object?> Y,
    IReadOnlyList<string>? Text,
    IReadOnlyList<double>? Size,
    ChartTraceStyle Style);

/// <summary>Per-series presentation. <see cref="SeriesType"/> is one of
/// <see cref="ChartSpecParser.SupportedSeriesTypes"/>.</summary>
public sealed record ChartTraceStyle(
    string SeriesType,
    string Name,
    bool SecondaryAxis,
    bool Visible,
    ChartLineStyle Line,
    ChartMarkerStyle Marker,
    string? FillColor,
    double? FillOpacity);

/// <summary>Series line presentation. <see cref="Width"/> null means
/// "Excel default".</summary>
public sealed record ChartLineStyle(
    string Color,
    string Dash,
    double? Width);

/// <summary>Series marker presentation. <see cref="Size"/> 0 means "no
/// markers".</summary>
public sealed record ChartMarkerStyle(
    double Size,
    string Color,
    string Shape);

/// <summary>A reference-line annotation: <c>event_line</c> (vertical, a
/// value on the x axis) or <c>threshold</c> (horizontal, a value on the
/// y axis). Rendered as an extra two-point series spanning the orthogonal
/// axis.</summary>
public sealed record ChartAnnotationSpec(
    int Id,
    string Type,
    string Axis,
    object? Value,
    ChartAnnotationStyle Style,
    string? Label);

/// <summary>Annotation line presentation. Null members mean "default".</summary>
public sealed record ChartAnnotationStyle(
    string? Color,
    double? Width,
    string Dash,
    double? Opacity);
