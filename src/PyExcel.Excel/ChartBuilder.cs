#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace PyExcel.Excel;

/// <summary>
/// Builds a native Excel chart from a validated <see cref="ChartSpecDocument"/>
/// — the v2 port of v1's <c>chartBuilder.bas</c>, late-bound over
/// <c>dynamic</c> COM (no Office PIA) like the rest of the host's COM
/// surface. Also embeds <see cref="ChartImage"/> payloads (Matplotlib
/// renders) as worksheet pictures.
///
/// <para><b>Orphan guard.</b> The ChartObject is created first and every
/// subsequent configuration step runs inside a try/catch that deletes it
/// before rethrowing — a failed build never leaves a half-configured
/// chart floating on the sheet (v1 routinely did).</para>
///
/// <para><b>Error policy.</b> Structural steps (adding the ChartObject,
/// adding a series, assigning its data) throw — without them there is no
/// chart to show. Cosmetic steps (colours, dashes, marker shapes, label
/// text) are individually guarded and logged to <see cref="Trace"/> on
/// failure: Excel rejects some combinations per chart type, and a missed
/// colour is not a reason to lose the chart. Nothing is silently
/// swallowed — every skipped cosmetic logs what failed.</para>
///
/// <para>Must run on Excel's main thread (COM). Callers queue via
/// <c>ExcelAsyncUtil.QueueAsMacro</c>.</para>
/// </summary>
public static class ChartBuilder
{
    // Excel / Office COM enum values (late-bound — no PIA to supply them).
    private const int XlArea = 1;
    private const int XlLine = 4;
    private const int XlPie = 5;
    private const int XlBubble = 15;
    private const int XlColumnClustered = 51;
    private const int XlColumnStacked = 52;
    private const int XlBarClustered = 57;
    private const int XlBarStacked = 58;
    private const int XlLineStacked = 63;
    private const int XlXYScatterLines = 74;
    private const int XlAreaStacked = 76;
    private const int XlXYScatter = -4169;

    private const int XlCategory = 1;
    private const int XlValue = 2;
    private const int XlSecondary = 2;

    private const int XlLegendPositionBottom = -4107;
    private const int XlLegendPositionRight = -4152;

    private const int XlScaleLinear = -4132;
    private const int XlScaleLogarithmic = -4133;

    private const int XlMarkerStyleSquare = 1;
    private const int XlMarkerStyleDiamond = 2;
    private const int XlMarkerStyleTriangle = 3;
    private const int XlMarkerStyleStar = 5;
    private const int XlMarkerStyleCircle = 8;
    private const int XlMarkerStylePlus = 9;
    private const int XlMarkerStyleNone = -4142;
    private const int XlMarkerStyleX = -4168;

    private const int MsoLineSolid = 1;
    private const int MsoLineRoundDot = 3;
    private const int MsoLineDash = 4;
    private const int MsoLineDashDot = 5;
    private const int MsoLineDashDotDot = 6;
    private const int MsoLineLongDash = 7;
    private const int MsoLineLongDashDot = 8;

    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    /// <summary>
    /// Build a chart on <paramref name="sheet"/> at the given position.
    /// On any failure the partially built ChartObject is deleted before
    /// the exception propagates.
    /// </summary>
    /// <param name="sheet">The target <c>Worksheet</c> (late-bound COM).</param>
    /// <param name="spec">A parsed, validated spec document.</param>
    public static void Build(
        dynamic sheet, ChartSpecDocument spec,
        double left, double top, double width, double height)
    {
        if (spec is null) throw new ArgumentNullException(nameof(spec));

        dynamic chartObject = sheet.ChartObjects().Add(left, top, width, height);
        try
        {
            dynamic chart = chartObject.Chart;
            chart.ChartType = MapChartType(spec.ChartType);

            if (spec.Title.Length > 0)
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = spec.Title;
            }
            else
            {
                chart.HasTitle = false;
            }

            var anyDates = false;
            var addedAny = false;
            foreach (var trace in spec.Traces)
            {
                var shaped = ChartTraceData.Shape(trace);
                if (shaped is null)
                {
                    Trace.WriteLine(
                        $"ChartBuilder: trace {trace.Id} ('{trace.Style.Name}') has no usable data points; skipped.");
                    continue;
                }
                AddSeries(chart, trace, shaped, spec.BarMode);
                anyDates |= shaped.XIsDate;
                addedAny = true;
            }
            if (!addedAny && spec.Traces.Count > 0)
                throw new FormatException(
                    "no trace in the chart spec produced any plottable data points");

            if (spec.ChartType != "pie")
            {
                ApplyAxis(chart, spec.XAxis, XlCategory, anyDates);
                ApplyAxis(chart, spec.YAxis, XlValue, dateTicks: false);
            }

            ApplyLegend(chart, spec.Legend);
            ApplyBarMode(chart, spec.BarMode);

            foreach (var annotation in spec.Annotations)
                AddAnnotation(chart, annotation);
        }
        catch
        {
            // Orphan guard: a failed build must not leave a half-configured
            // chart on the sheet. Best-effort — the original exception is
            // what the user needs to see, not a secondary delete failure.
            try { chartObject.Delete(); }
            catch (Exception cleanupEx)
            {
                Trace.WriteLine($"ChartBuilder: orphan cleanup failed — {cleanupEx.Message}");
            }
            throw;
        }
    }

    /// <summary>
    /// Embed a rendered figure image as a worksheet picture at the given
    /// position. The bytes are staged through a temp file (Shapes.AddPicture
    /// only takes a path); the file is deleted after the insert since the
    /// picture is stored with the workbook.
    /// </summary>
    public static void EmbedImage(dynamic sheet, ChartImage image, double left, double top)
    {
        if (image is null) throw new ArgumentNullException(nameof(image));

        var tempPath = Path.Combine(
            Path.GetTempPath(),
            $"pyexcel_chart_{Guid.NewGuid():N}.{image.Format}");
        File.WriteAllBytes(tempPath, image.Data);
        try
        {
            // LinkToFile=false, SaveWithDocument=true; -1 width/height keeps
            // the image's native size.
            sheet.Shapes.AddPicture(tempPath, MsoFalse, MsoTrue, left, top, -1f, -1f);
        }
        finally
        {
            try { File.Delete(tempPath); }
            catch (IOException ioEx)
            {
                Trace.WriteLine($"ChartBuilder: temp image cleanup failed — {ioEx.Message}");
            }
        }
    }

    // -------------------------------------------------------------------------
    // Series
    // -------------------------------------------------------------------------

    private static void AddSeries(
        dynamic chart, ChartTraceSpec trace, ChartTraceData.Shaped shaped, string barMode)
    {
        dynamic series = chart.SeriesCollection().NewSeries();
        series.Name = trace.Style.Name;
        series.Values = shaped.YValues;
        if (shaped.XValues is { } xs)
            series.XValues = xs;

        series.ChartType = MapSeriesType(trace.Style.SeriesType, barMode);

        if (trace.Style.SecondaryAxis)
            Cosmetic(() => series.AxisGroup = XlSecondary,
                $"trace {trace.Id}: secondary axis assignment");

        if (!trace.Style.Visible)
            HideSeries(series, trace.Id);

        if (trace.Style.SeriesType == "bubble" && trace.Size is { Count: > 0 } sizes)
        {
            var sizeArray = new object[sizes.Count];
            for (var i = 0; i < sizes.Count; i++) sizeArray[i] = sizes[i];
            Cosmetic(() => series.BubbleSizes = sizeArray,
                $"trace {trace.Id}: bubble sizes");
        }

        ApplyLineStyle(series, trace);
        ApplyMarkerStyle(series, trace);
        ApplyFillStyle(series, trace);

        // Histogram text rides as category labels (already in XValues);
        // every other type renders text as per-point data labels.
        if (trace.Style.SeriesType != "histogram" && trace.Text is { Count: > 0 } text)
            ApplyTextLabels(series, text, trace.Id);
    }

    private static void HideSeries(dynamic series, int traceId)
    {
        try
        {
            // Excel 2013+: filtered out of the plot but kept in the data.
            series.IsFiltered = true;
        }
        catch (Exception ex)
        {
            Trace.WriteLine(
                $"ChartBuilder: trace {traceId}: IsFiltered unavailable ({ex.Message}); hiding visuals instead.");
            Cosmetic(() => series.MarkerStyle = XlMarkerStyleNone, $"trace {traceId}: hide markers");
            Cosmetic(() => series.Format.Line.Visible = MsoFalse, $"trace {traceId}: hide line");
            Cosmetic(() => series.Format.Fill.Visible = MsoFalse, $"trace {traceId}: hide fill");
        }
    }

    private static void ApplyLineStyle(dynamic series, ChartTraceSpec trace)
    {
        var hasLine = trace.Style.SeriesType switch
        {
            "scatter_lines" or "scatter_lines_markers" or "line"
                or "area" or "bar" or "column" => true,
            _ => false,
        };
        if (!hasLine)
        {
            // Marker-only scatter: no connecting line.
            if (trace.Style.SeriesType is "scatter" or "bubble")
                Cosmetic(() => series.Format.Line.Visible = MsoFalse,
                    $"trace {trace.Id}: line off");
            return;
        }

        var width = trace.Style.Line.Width is { } w && w > 0 ? w : 1.0;
        Cosmetic(
            () =>
            {
                dynamic line = series.Format.Line;
                line.Visible = MsoTrue;
                line.ForeColor.RGB = ChartColor.Parse(trace.Style.Line.Color);
                line.Weight = (float)width;
                line.DashStyle = MapDash(trace.Style.Line.Dash);
            },
            $"trace {trace.Id}: line styling");
    }

    private static void ApplyMarkerStyle(dynamic series, ChartTraceSpec trace)
    {
        var supportsMarkers = trace.Style.SeriesType switch
        {
            "scatter" or "scatter_lines" or "scatter_lines_markers" or "line" or "bubble" => true,
            _ => false,
        };
        if (!supportsMarkers) return;

        if (trace.Style.Marker.Size <= 0)
        {
            Cosmetic(() => series.MarkerStyle = XlMarkerStyleNone,
                $"trace {trace.Id}: markers off");
            return;
        }

        Cosmetic(
            () =>
            {
                series.MarkerStyle = MapMarkerShape(trace.Style.Marker.Shape);
                series.MarkerSize = (int)Math.Round(trace.Style.Marker.Size);
                series.MarkerForegroundColor = ChartColor.Parse(trace.Style.Marker.Color);
                series.MarkerBackgroundColor = ChartColor.Parse(trace.Style.Marker.Color);
            },
            $"trace {trace.Id}: marker styling");
    }

    private static void ApplyFillStyle(dynamic series, ChartTraceSpec trace)
    {
        if (trace.Style.FillColor is null && trace.Style.FillOpacity is null) return;

        Cosmetic(
            () =>
            {
                dynamic fill = series.Format.Fill;
                fill.Visible = MsoTrue;
                if (trace.Style.FillColor is { } color)
                    fill.ForeColor.RGB = ChartColor.Parse(color);
                if (trace.Style.FillOpacity is { } opacity)
                {
                    // Spec carries opacity (1 = opaque); COM wants transparency.
                    var t = 1.0 - Math.Max(0.0, Math.Min(1.0, opacity));
                    fill.Transparency = (float)t;
                }
            },
            $"trace {trace.Id}: fill styling");
    }

    private static void ApplyTextLabels(
        dynamic series, IReadOnlyList<string> text, int traceId)
    {
        Cosmetic(
            () =>
            {
                series.ApplyDataLabels();
                dynamic labels = series.DataLabels();
                labels.ShowValue = false;
                labels.ShowSeriesName = false;
                labels.ShowCategoryName = false;

                int points = series.Points().Count;
                var n = Math.Min(points, text.Count);
                for (var i = 1; i <= n; i++)
                {
                    dynamic point = series.Points(i);
                    point.HasDataLabel = true;
                    point.DataLabel.Text = text[i - 1];
                }
            },
            $"trace {traceId}: per-point text labels");
    }

    // -------------------------------------------------------------------------
    // Axes / legend / bar mode / annotations
    // -------------------------------------------------------------------------

    private static void ApplyAxis(dynamic chart, ChartAxisSpec spec, int axisType, bool dateTicks)
    {
        dynamic axis = chart.Axes(axisType);

        if (spec.Title.Length > 0)
        {
            axis.HasTitle = true;
            axis.AxisTitle.Text = spec.Title;
        }
        else
        {
            axis.HasTitle = false;
        }

        if (dateTicks)
            Cosmetic(() => axis.TickLabels.NumberFormat = "yyyy-mm-dd",
                "x axis: date tick format");

        if (spec.LogScale)
        {
            // Excel rejects log scaling in some configurations (non-positive
            // data); fall back to linear rather than failing the build.
            try
            {
                axis.ScaleType = XlScaleLogarithmic;
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"ChartBuilder: log scale rejected ({ex.Message}); using linear.");
                axis.ScaleType = XlScaleLinear;
            }
        }
        else
        {
            axis.ScaleType = XlScaleLinear;
        }

        // Log axes only accept positive explicit bounds.
        var isLog = spec.LogScale;
        if (spec.Min is { } min && (!isLog || min > 0))
            axis.MinimumScale = min;
        else
            axis.MinimumScaleIsAuto = true;
        if (spec.Max is { } max && (!isLog || max > 0))
            axis.MaximumScale = max;
        else
            axis.MaximumScaleIsAuto = true;
    }

    private static void ApplyLegend(dynamic chart, ChartLegendSpec legend)
    {
        chart.HasLegend = legend.Visible;
        if (!legend.Visible) return;
        chart.Legend.Position = legend.Position == "bottom"
            ? XlLegendPositionBottom
            : XlLegendPositionRight;
    }

    private static void ApplyBarMode(dynamic chart, string barMode)
    {
        // "stack" is handled per-series in MapSeriesType; "overlay" maps to
        // full overlap with translucent fills (the Plotly look); "group" is
        // Excel's default clustering.
        if (barMode != "overlay") return;
        Cosmetic(
            () =>
            {
                if ((int)chart.ChartGroups().Count > 0)
                {
                    dynamic group = chart.ChartGroups(1);
                    group.Overlap = 100;
                    group.GapWidth = 10;
                }
            },
            "barmode overlay");
    }

    private static void AddAnnotation(dynamic chart, ChartAnnotationSpec annotation)
    {
        if (!TryAnnotationValue(annotation.Value, out var value))
        {
            Trace.WriteLine(
                $"ChartBuilder: annotation {annotation.Id} has a non-numeric value " +
                $"({annotation.Value}); skipped.");
            return;
        }

        // The line spans the orthogonal axis's current scale, read after the
        // data series landed so auto-scaling has settled.
        var vertical = annotation.Type == "event_line";
        dynamic spanAxis = chart.Axes(vertical ? XlValue : XlCategory);
        double spanMin = spanAxis.MinimumScale;
        double spanMax = spanAxis.MaximumScale;

        dynamic series = chart.SeriesCollection().NewSeries();
        series.Name = annotation.Label
            ?? (vertical ? $"Event {annotation.Id}" : $"Threshold {annotation.Id}");
        if (vertical)
        {
            series.XValues = new object[] { value, value };
            series.Values = new object[] { spanMin, spanMax };
        }
        else
        {
            series.XValues = new object[] { spanMin, spanMax };
            series.Values = new object[] { value, value };
        }
        series.ChartType = XlXYScatterLines;
        Cosmetic(() => series.MarkerStyle = XlMarkerStyleNone,
            $"annotation {annotation.Id}: markers off");

        Cosmetic(
            () =>
            {
                dynamic line = series.Format.Line;
                line.Visible = MsoTrue;
                if (annotation.Style.Color is { } color)
                    line.ForeColor.RGB = ChartColor.Parse(color);
                line.Weight = (float)(annotation.Style.Width is { } w && w > 0 ? w : 2.0);
                line.DashStyle = MapDash(annotation.Style.Dash);
                if (annotation.Style.Opacity is { } opacity)
                    line.Transparency = (float)(1.0 - Math.Max(0.0, Math.Min(1.0, opacity)));
            },
            $"annotation {annotation.Id}: line styling");
    }

    private static bool TryAnnotationValue(object? raw, out double value)
    {
        switch (raw)
        {
            case double d:
                value = d;
                return true;
            case long l:
                value = l;
                return true;
            case string s when DateTime.TryParseExact(
                s, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None, out var date):
                value = date.ToOADate();
                return true;
            default:
                value = 0;
                return false;
        }
    }

    // -------------------------------------------------------------------------
    // Enum mapping
    // -------------------------------------------------------------------------

    private static int MapChartType(string chartType) => chartType switch
    {
        "bar" => XlBarClustered,
        "line" => XlLine,
        "area" => XlArea,
        "pie" => XlPie,
        _ => XlXYScatter, // "xy" — the parser rejects anything else
    };

    private static int MapSeriesType(string seriesType, string barMode)
    {
        var stacked = barMode == "stack";
        return seriesType switch
        {
            "scatter" => XlXYScatter,
            "scatter_lines" or "scatter_lines_markers" => XlXYScatterLines,
            "line" => stacked ? XlLineStacked : XlLine,
            "column" or "histogram" => stacked ? XlColumnStacked : XlColumnClustered,
            "bar" => stacked ? XlBarStacked : XlBarClustered,
            "area" => stacked ? XlAreaStacked : XlArea,
            "pie" => XlPie,
            "bubble" => XlBubble,
            // The parser guarantees the supported set; this arm is the
            // belt-and-braces for a model constructed without it.
            _ => throw new FormatException($"unsupported series type '{seriesType}'"),
        };
    }

    private static int MapDash(string dash) => dash switch
    {
        "dash" => MsoLineDash,
        "dot" => MsoLineRoundDot,
        "dashdot" => MsoLineDashDot,
        "dashdotdot" => MsoLineDashDotDot,
        "longdash" => MsoLineLongDash,
        "longdashdot" => MsoLineLongDashDot,
        _ => MsoLineSolid,
    };

    private static int MapMarkerShape(string shape) => shape switch
    {
        "square" => XlMarkerStyleSquare,
        "diamond" => XlMarkerStyleDiamond,
        "triangle" => XlMarkerStyleTriangle,
        "x" => XlMarkerStyleX,
        "plus" => XlMarkerStylePlus,
        "star" => XlMarkerStyleStar,
        "none" => XlMarkerStyleNone,
        _ => XlMarkerStyleCircle,
    };

    /// <summary>Run one cosmetic configuration step; on failure log what
    /// was skipped and continue. Excel rejects some styling calls
    /// depending on chart type / version — a missed colour or label must
    /// not cost the user the chart.</summary>
    private static void Cosmetic(Action step, string description)
    {
        try
        {
            step();
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"ChartBuilder: skipped {description} — {ex.Message}");
        }
    }
}
#endif
