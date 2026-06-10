using System;
using System.Collections.Generic;
using System.Globalization;

namespace PyExcel.Excel;

/// <summary>
/// Shapes one trace's raw spec data (<c>x</c> / <c>y</c> JSON cells) into
/// the arrays the COM builder assigns to <c>Series.XValues</c> /
/// <c>Series.Values</c>. This is the cross-platform port of v1
/// <c>chartBuilder.bas</c>'s <c>BuildXYForCategories</c> /
/// <c>BuildXYNumericFlexible</c>, kept COM-free so Linux CI covers the
/// rules:
///
/// <list type="bullet">
///   <item>Categorical series (bar / column / pie / histogram): x cells
///     become category strings, y must be numeric; pairs with a
///     non-numeric y are dropped.</item>
///   <item>Numeric series: if every x cell is a number or an ISO-8601
///     date, x stays numeric (dates convert to OADate serials and the
///     result is flagged so the builder applies a date number format);
///     one non-numeric, non-date x cell flips the whole axis to
///     categories. Pairs with a null/non-numeric member are dropped.</item>
///   <item>Histogram traces prefer the pre-binned <c>text</c> labels as
///     categories (the kernel emits <c>lo:hi</c> range labels) over the
///     raw bin midpoints.</item>
/// </list>
/// </summary>
public static class ChartTraceData
{
    /// <summary>The shaped, COM-ready data for one series.
    /// <see cref="XValues"/> is null when the trace carried no usable x
    /// (Excel then numbers the points 1..N). <see cref="XIsDate"/> is
    /// true when x values were ISO dates converted to OADate serials, so
    /// the builder can format the axis ticks as dates.</summary>
    public sealed record Shaped(
        object[]? XValues,
        object[] YValues,
        bool XIsDate);

    /// <summary>Whether a series type plots y values against category
    /// labels (vs. a numeric/date x axis).</summary>
    public static bool IsCategorical(string seriesType) => seriesType switch
    {
        "bar" or "column" or "pie" or "histogram" => true,
        _ => false,
    };

    /// <summary>Shape one trace. Returns null when no usable data points
    /// remain after filtering — the builder skips the series entirely
    /// (matching v1's "skip trace" behaviour) rather than adding an empty
    /// series Excel would render as garbage.</summary>
    public static Shaped? Shape(ChartTraceSpec trace)
    {
        if (trace is null) throw new ArgumentNullException(nameof(trace));

        var x = trace.X;
        // Histogram: the kernel pre-bins and ships human-readable range
        // labels in text; those are the categories the user should see,
        // not the numeric bin midpoints.
        if (trace.Style.SeriesType == "histogram"
            && trace.Text is { } labels && labels.Count == trace.Y.Count)
        {
            var asObjects = new object?[labels.Count];
            for (var i = 0; i < labels.Count; i++) asObjects[i] = labels[i];
            x = asObjects;
        }

        return IsCategorical(trace.Style.SeriesType)
            ? ShapeCategorical(x, trace.Y)
            : ShapeNumeric(x, trace.Y);
    }

    private static Shaped? ShapeCategorical(
        IReadOnlyList<object?> x, IReadOnlyList<object?> y)
    {
        var n = Math.Min(x.Count == 0 ? int.MaxValue : x.Count, y.Count);
        if (y.Count == 0) return null;

        var xs = new List<object>(y.Count);
        var ys = new List<object>(y.Count);
        for (var i = 0; i < n; i++)
        {
            if (!TryToDouble(y[i], out var yv)) continue;
            xs.Add(CategoryLabel(x.Count == 0 ? null : x[i]));
            ys.Add(yv);
        }
        if (ys.Count == 0) return null;
        return new Shaped(
            XValues: x.Count == 0 ? null : xs.ToArray(),
            YValues: ys.ToArray(),
            XIsDate: false);
    }

    private static Shaped? ShapeNumeric(
        IReadOnlyList<object?> x, IReadOnlyList<object?> y)
    {
        if (y.Count == 0) return null;

        if (x.Count == 0)
        {
            // No x at all: y-only series, Excel numbers the points.
            var onlyY = new List<object>(y.Count);
            foreach (var cell in y)
                if (TryToDouble(cell, out var yv)) onlyY.Add(yv);
            return onlyY.Count == 0 ? null : new Shaped(null, onlyY.ToArray(), false);
        }

        // One pass to classify the x axis: numeric, date, or category.
        var sawDate = false;
        var treatAsCategory = false;
        foreach (var cell in x)
        {
            if (cell is null) continue;
            if (IsJsonNumber(cell)) continue;
            if (cell is string s && TryParseIsoDate(s, out _)) { sawDate = true; continue; }
            treatAsCategory = true;
            break;
        }

        var n = Math.Min(x.Count, y.Count);
        var xs = new List<object>(n);
        var ys = new List<object>(n);
        for (var i = 0; i < n; i++)
        {
            if (x[i] is null || !TryToDouble(y[i], out var yv)) continue;

            if (treatAsCategory)
            {
                xs.Add(CategoryLabel(x[i]));
            }
            else if (x[i] is string s)
            {
                if (!TryParseIsoDate(s, out var date)) continue;
                xs.Add(date.ToOADate());
            }
            else
            {
                if (!TryToDouble(x[i], out var xv)) continue;
                xs.Add(xv);
            }
            ys.Add(yv);
        }
        if (ys.Count == 0) return null;
        return new Shaped(xs.ToArray(), ys.ToArray(), XIsDate: sawDate && !treatAsCategory);
    }

    private static string CategoryLabel(object? cell) => cell switch
    {
        null => string.Empty,
        string s => s,
        double d => d.ToString("R", CultureInfo.InvariantCulture),
        long l => l.ToString(CultureInfo.InvariantCulture),
        bool b => b ? "TRUE" : "FALSE",
        _ => Convert.ToString(cell, CultureInfo.InvariantCulture) ?? string.Empty,
    };

    private static bool IsJsonNumber(object cell) => cell is double or long;

    private static bool TryToDouble(object? cell, out double value)
    {
        switch (cell)
        {
            case double d when !double.IsNaN(d):
                value = d;
                return true;
            case long l:
                value = l;
                return true;
            default:
                value = 0;
                return false;
        }
    }

    /// <summary>ISO-8601 date or datetime (<c>yyyy-MM-dd</c>, optionally
    /// with a <c>T</c>-separated time part) — the only string forms the
    /// kernel emits for date-typed figure data.</summary>
    private static bool TryParseIsoDate(string s, out DateTime date)
    {
        var formats = new[]
        {
            "yyyy-MM-dd",
            "yyyy-MM-dd'T'HH:mm:ss",
            "yyyy-MM-dd'T'HH:mm:ss.FFFFFF",
        };
        return DateTime.TryParseExact(
            s, formats, CultureInfo.InvariantCulture,
            DateTimeStyles.None, out date);
    }
}
