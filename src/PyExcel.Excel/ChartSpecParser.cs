using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using PyExcel.Bridge;

namespace PyExcel.Excel;

/// <summary>
/// Parses + validates the chart-spec JSON the kernel emits for a Plotly
/// figure (see <c>embedded/pyexcel/kernel/chart.py</c> for the schema)
/// into the typed <see cref="ChartSpecDocument"/> model.
///
/// <para>This is the cross-platform half of the chart pipeline — pure
/// logic, no COM — so the spec contract is testable on Linux CI. The
/// net48-only <c>ChartBuilder</c> consumes the validated model and never
/// has to defend against a malformed document.</para>
///
/// <para><b>Validation policy.</b> Structural problems (wrong version,
/// non-object root, unknown series type, duplicate trace ids, mismatched
/// x/y lengths) throw <see cref="FormatException"/> with a message that
/// names the offending element — these reach the user via LogDisplay, so
/// they must be actionable. Merely <em>missing optional</em> attributes
/// fall back to safe defaults (empty title, auto-scaled axes, visible
/// legend), matching the producing side's own defaults.</para>
/// </summary>
public static class ChartSpecParser
{
    /// <summary>The schema version this parser understands.</summary>
    public const int SupportedVersion = 1;

    /// <summary>Chart-level base types the builder renders.</summary>
    public static readonly IReadOnlyList<string> SupportedChartTypes =
        new[] { "xy", "line", "bar", "area", "pie" };

    /// <summary>Per-series types the builder renders. Mirrors the Python
    /// side's <c>SUPPORTED_SERIES_TYPES</c>.</summary>
    public static readonly IReadOnlyList<string> SupportedSeriesTypes =
        new[]
        {
            "scatter", "scatter_lines", "scatter_lines_markers", "line",
            "column", "bar", "area", "pie", "histogram", "bubble",
        };

    /// <summary>Annotation types the builder renders.</summary>
    public static readonly IReadOnlyList<string> SupportedAnnotationTypes =
        new[] { "event_line", "threshold" };

    /// <summary>
    /// Parse a chart-spec JSON document. Throws <see cref="FormatException"/>
    /// for anything the builder could not render faithfully.
    /// </summary>
    public static ChartSpecDocument Parse(string json)
    {
        if (json is null) throw new ArgumentNullException(nameof(json));

        object? root;
        try
        {
            root = CanonicalJson.Decode(Encoding.UTF8.GetBytes(json));
        }
        catch (FormatException exc)
        {
            throw new FormatException($"chart spec is not valid JSON: {exc.Message}", exc);
        }

        if (root is not Dictionary<string, object?> doc)
            throw new FormatException("chart spec root must be a JSON object");

        var version = GetRequiredInt(doc, "version", "chart spec");
        if (version != SupportedVersion)
            throw new FormatException(
                $"unsupported chart spec version {version}; this host understands version {SupportedVersion}");

        var chartType = GetString(doc, "chart_type") ?? "xy";
        if (!Contains(SupportedChartTypes, chartType))
            throw new FormatException(
                $"unsupported chart type '{chartType}'; supported: {string.Join(", ", SupportedChartTypes)}");

        var traces = ParseTraces(doc);
        var annotations = ParseAnnotations(doc);

        return new ChartSpecDocument(
            Version: version,
            ChartType: chartType,
            Title: GetString(doc, "title") ?? string.Empty,
            XAxis: ParseAxis(GetObject(doc, "x_axis"), "x_axis"),
            YAxis: ParseAxis(GetObject(doc, "y_axis"), "y_axis"),
            Legend: ParseLegend(GetObject(doc, "legend")),
            BarMode: GetString(doc, "barmode") ?? "group",
            Traces: traces,
            Annotations: annotations);
    }

    // -------------------------------------------------------------------------
    // Section parsers
    // -------------------------------------------------------------------------

    private static ChartAxisSpec ParseAxis(Dictionary<string, object?>? axis, string label)
    {
        if (axis is null)
            return new ChartAxisSpec(string.Empty, null, null, false);
        return new ChartAxisSpec(
            Title: GetString(axis, "title") ?? string.Empty,
            Min: GetNumber(axis, "min", label),
            Max: GetNumber(axis, "max", label),
            LogScale: GetBool(axis, "log_scale") ?? false);
    }

    private static ChartLegendSpec ParseLegend(Dictionary<string, object?>? legend)
    {
        if (legend is null)
            return new ChartLegendSpec(Visible: true, Position: "right");
        var position = GetString(legend, "position") ?? "right";
        if (position != "right" && position != "bottom")
            throw new FormatException(
                $"legend position must be 'right' or 'bottom', got '{position}'");
        return new ChartLegendSpec(
            Visible: GetBool(legend, "visible") ?? true,
            Position: position);
    }

    private static IReadOnlyList<ChartTraceSpec> ParseTraces(Dictionary<string, object?> doc)
    {
        if (!doc.TryGetValue("traces", out var raw) || raw is null)
            return Array.Empty<ChartTraceSpec>();
        if (raw is not List<object?> list)
            throw new FormatException("'traces' must be a JSON array");

        var traces = new List<ChartTraceSpec>(list.Count);
        var seenIds = new HashSet<int>();
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i] is not Dictionary<string, object?> traceObj)
                throw new FormatException($"trace at index {i} must be a JSON object");
            var trace = ParseTrace(traceObj, i);
            if (!seenIds.Add(trace.Id))
                throw new FormatException($"duplicate trace id {trace.Id}");
            traces.Add(trace);
        }
        return traces;
    }

    private static ChartTraceSpec ParseTrace(Dictionary<string, object?> trace, int index)
    {
        var context = $"trace at index {index}";
        var id = GetRequiredInt(trace, "id", context);

        var x = GetValueList(trace, "x") ?? Array.Empty<object?>();
        var y = GetValueList(trace, "y") ?? Array.Empty<object?>();
        if (x.Count > 0 && y.Count > 0 && x.Count != y.Count)
            throw new FormatException(
                $"trace {id}: x has {x.Count} values but y has {y.Count}");

        var style = ParseStyle(GetObject(trace, "style"), id);

        return new ChartTraceSpec(
            Id: id,
            X: x,
            Y: y,
            Text: GetStringList(trace, "text"),
            Size: GetNumberList(trace, "size", $"trace {id} size"),
            Style: style);
    }

    private static ChartTraceStyle ParseStyle(Dictionary<string, object?>? style, int traceId)
    {
        if (style is null)
            throw new FormatException($"trace {traceId} is missing its 'style' object");

        var seriesType = GetString(style, "series_type")
            ?? throw new FormatException($"trace {traceId} style is missing 'series_type'");
        if (!Contains(SupportedSeriesTypes, seriesType))
            throw new FormatException(
                $"trace {traceId}: unsupported series type '{seriesType}'; " +
                $"supported: {string.Join(", ", SupportedSeriesTypes)}");

        var line = GetObject(style, "line");
        var marker = GetObject(style, "marker");
        var axisGroup = GetString(style, "axis_group") ?? "primary";

        return new ChartTraceStyle(
            SeriesType: seriesType,
            Name: GetString(style, "name") ?? $"Series {traceId}",
            SecondaryAxis: string.Equals(axisGroup, "secondary", StringComparison.Ordinal),
            Visible: GetBool(style, "visible") ?? true,
            Line: new ChartLineStyle(
                Color: line is null ? "#000000" : GetString(line, "color") ?? "#000000",
                Dash: line is null ? "solid" : GetString(line, "dash") ?? "solid",
                Width: line is null ? null : GetNumber(line, "width", $"trace {traceId} line")),
            Marker: new ChartMarkerStyle(
                Size: marker is null ? 6.0 : GetNumber(marker, "size", $"trace {traceId} marker") ?? 6.0,
                Color: marker is null ? "#000000" : GetString(marker, "color") ?? "#000000",
                Shape: marker is null ? "circle" : GetString(marker, "shape") ?? "circle"),
            FillColor: GetString(style, "fill_color"),
            FillOpacity: GetNumber(style, "fill_opacity", $"trace {traceId} style"));
    }

    private static IReadOnlyList<ChartAnnotationSpec> ParseAnnotations(Dictionary<string, object?> doc)
    {
        if (!doc.TryGetValue("annotations", out var raw) || raw is null)
            return Array.Empty<ChartAnnotationSpec>();
        if (raw is not List<object?> list)
            throw new FormatException("'annotations' must be a JSON array");

        var annotations = new List<ChartAnnotationSpec>(list.Count);
        for (var i = 0; i < list.Count; i++)
        {
            if (list[i] is not Dictionary<string, object?> annObj)
                throw new FormatException($"annotation at index {i} must be a JSON object");

            var context = $"annotation at index {i}";
            var id = GetRequiredInt(annObj, "id", context);
            var type = GetString(annObj, "type")
                ?? throw new FormatException($"{context} is missing 'type'");
            if (!Contains(SupportedAnnotationTypes, type))
                throw new FormatException(
                    $"annotation {id}: unsupported type '{type}'; " +
                    $"supported: {string.Join(", ", SupportedAnnotationTypes)}");
            var axis = GetString(annObj, "axis")
                ?? throw new FormatException($"annotation {id} is missing 'axis'");
            if (axis != "x" && axis != "y")
                throw new FormatException($"annotation {id}: axis must be 'x' or 'y', got '{axis}'");
            if (!annObj.TryGetValue("value", out var value) || value is null)
                throw new FormatException($"annotation {id} is missing 'value'");

            var style = GetObject(annObj, "style");
            annotations.Add(new ChartAnnotationSpec(
                Id: id,
                Type: type,
                Axis: axis,
                Value: value,
                Style: new ChartAnnotationStyle(
                    Color: style is null ? null : GetString(style, "color"),
                    Width: style is null ? null : GetNumber(style, "width", $"annotation {id} style"),
                    Dash: style is null ? "solid" : GetString(style, "dash") ?? "solid",
                    Opacity: style is null ? null : GetNumber(style, "opacity", $"annotation {id} style")),
                Label: GetString(annObj, "label")));
        }
        return annotations;
    }

    // -------------------------------------------------------------------------
    // Typed accessors over CanonicalJson's object model
    // (Dictionary / List / string / long / double / bool / null)
    // -------------------------------------------------------------------------

    private static bool Contains(IReadOnlyList<string> set, string value)
    {
        for (var i = 0; i < set.Count; i++)
            if (string.Equals(set[i], value, StringComparison.Ordinal)) return true;
        return false;
    }

    private static Dictionary<string, object?>? GetObject(
        Dictionary<string, object?> obj, string key)
        => obj.TryGetValue(key, out var v) && v is Dictionary<string, object?> d ? d : null;

    private static string? GetString(Dictionary<string, object?> obj, string key)
        => obj.TryGetValue(key, out var v) && v is string s ? s : null;

    private static bool? GetBool(Dictionary<string, object?> obj, string key)
        => obj.TryGetValue(key, out var v) && v is bool b ? b : (bool?)null;

    private static double? GetNumber(
        Dictionary<string, object?> obj, string key, string context)
    {
        if (!obj.TryGetValue(key, out var v) || v is null) return null;
        return v switch
        {
            double d => d,
            long l => l,
            _ => throw new FormatException(
                $"{context}: '{key}' must be a number, got {Describe(v)}"),
        };
    }

    private static int GetRequiredInt(
        Dictionary<string, object?> obj, string key, string context)
    {
        if (!obj.TryGetValue(key, out var v) || v is null)
            throw new FormatException($"{context} is missing required '{key}'");
        return v switch
        {
            long l => checked((int)l),
            _ => throw new FormatException(
                $"{context}: '{key}' must be an integer, got {Describe(v)}"),
        };
    }

    private static IReadOnlyList<object?>? GetValueList(
        Dictionary<string, object?> obj, string key)
        => obj.TryGetValue(key, out var v) && v is List<object?> list ? list : null;

    private static IReadOnlyList<string>? GetStringList(
        Dictionary<string, object?> obj, string key)
    {
        if (!obj.TryGetValue(key, out var v) || v is not List<object?> list) return null;
        var result = new string[list.Count];
        for (var i = 0; i < list.Count; i++)
            result[i] = list[i] is null
                ? string.Empty
                : Convert.ToString(list[i], CultureInfo.InvariantCulture) ?? string.Empty;
        return result;
    }

    private static IReadOnlyList<double>? GetNumberList(
        Dictionary<string, object?> obj, string key, string context)
    {
        if (!obj.TryGetValue(key, out var v) || v is not List<object?> list) return null;
        var result = new double[list.Count];
        for (var i = 0; i < list.Count; i++)
        {
            result[i] = list[i] switch
            {
                double d => d,
                long l => l,
                _ => throw new FormatException(
                    $"{context}: entry {i} must be a number, got {Describe(list[i])}"),
            };
        }
        return result;
    }

    private static string Describe(object? value) => value switch
    {
        null => "null",
        string => "a string",
        bool => "a boolean",
        List<object?> => "an array",
        Dictionary<string, object?> => "an object",
        _ => value.GetType().Name,
    };
}
