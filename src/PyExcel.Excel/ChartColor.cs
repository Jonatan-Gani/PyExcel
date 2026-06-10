using System;
using System.Globalization;

namespace PyExcel.Excel;

/// <summary>
/// Parses the colour strings a chart spec carries into the OLE colour
/// integer Excel's COM surface expects (<c>R + G·256 + B·65536</c> — the
/// same packing as VBA's <c>RGB()</c>).
///
/// <para>Accepted forms — the union of what Plotly emits and what v1's
/// <c>chartBuilder.bas</c> accepted: <c>#RRGGBB</c>, <c>#RGB</c>,
/// <c>rgb(r,g,b)</c>, <c>rgba(r,g,b,a)</c> (alpha carried separately as
/// fill opacity, so it's ignored here), and a small set of CSS colour
/// names. Anything unrecognised parses as black rather than throwing —
/// a wrong colour is a cosmetic miss, not a reason to lose the chart.</para>
/// </summary>
public static class ChartColor
{
    /// <summary>Parse a colour spec to an OLE BGR-packed integer.
    /// Null/blank/unrecognised input parses as black (0).</summary>
    public static int Parse(string? colorSpec)
    {
        if (string.IsNullOrWhiteSpace(colorSpec)) return Pack(0, 0, 0);
        var spec = colorSpec!.Trim();

        if (spec[0] == '#')
            return ParseHex(spec);

        if (spec.StartsWith("rgba(", StringComparison.OrdinalIgnoreCase) && spec.EndsWith(")", StringComparison.Ordinal))
            return ParseRgbList(spec.Substring(5, spec.Length - 6), expectAlpha: true);

        if (spec.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase) && spec.EndsWith(")", StringComparison.Ordinal))
            return ParseRgbList(spec.Substring(4, spec.Length - 5), expectAlpha: false);

        return ParseNamed(spec);
    }

    private static int ParseHex(string spec)
    {
        var hex = spec.Substring(1);
        if (hex.Length == 3)
        {
            // #RGB → #RRGGBB
            hex = new string(new[] { hex[0], hex[0], hex[1], hex[1], hex[2], hex[2] });
        }
        if (hex.Length != 6
            || !int.TryParse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var r)
            || !int.TryParse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var g)
            || !int.TryParse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var b))
        {
            return Pack(0, 0, 0);
        }
        return Pack(r, g, b);
    }

    private static int ParseRgbList(string inner, bool expectAlpha)
    {
        var parts = inner.Split(',');
        var expected = expectAlpha ? 4 : 3;
        if (parts.Length != expected) return Pack(0, 0, 0);

        // Channels may be ints or (from some Plotly themes) floats.
        if (!double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var r)
            || !double.TryParse(parts[1].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var g)
            || !double.TryParse(parts[2].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var b))
        {
            return Pack(0, 0, 0);
        }
        return Pack(Clamp(r), Clamp(g), Clamp(b));
    }

    private static int ParseNamed(string spec) => spec.ToLowerInvariant() switch
    {
        "red" => Pack(255, 0, 0),
        "green" => Pack(0, 128, 0),
        "blue" => Pack(0, 0, 255),
        "yellow" => Pack(255, 255, 0),
        "cyan" => Pack(0, 255, 255),
        "magenta" => Pack(255, 0, 255),
        "white" => Pack(255, 255, 255),
        "gray" or "grey" => Pack(128, 128, 128),
        "orange" => Pack(255, 165, 0),
        "purple" => Pack(128, 0, 128),
        "black" => Pack(0, 0, 0),
        _ => Pack(0, 0, 0),
    };

    private static int Clamp(double v)
    {
        if (v < 0) return 0;
        if (v > 255) return 255;
        return (int)Math.Round(v, MidpointRounding.AwayFromZero);
    }

    private static int Pack(int r, int g, int b) => r | (g << 8) | (b << 16);
}
