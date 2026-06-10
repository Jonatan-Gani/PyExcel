using System;

namespace PyExcel.Excel;

/// <summary>
/// A rendered figure image produced by the kernel when a user
/// <c>transform()</c> returns a Matplotlib figure — SVG when the figure
/// could render to vector, PNG otherwise. The host embeds it as a
/// worksheet picture.
///
/// <para>The wire representation is a 1×1 Arrow binary column whose
/// schema carries <c>pyexcel-shape = image</c> and whose field carries
/// <c>pyexcel-image-format = svg|png</c>. See <see cref="ArrowMarshal"/>
/// for the decode plumbing and <c>embedded/pyexcel/kernel/chart.py</c>
/// for the producing side.</para>
/// </summary>
public sealed class ChartImage
{
    /// <summary>Image formats the kernel may emit.</summary>
    public const string FormatSvg = "svg";
    public const string FormatPng = "png";

    /// <summary>Raw image bytes.</summary>
    public byte[] Data { get; }

    /// <summary>Rendered format: <see cref="FormatSvg"/> or
    /// <see cref="FormatPng"/>. Doubles as the file extension when the
    /// host stages the image for embedding.</summary>
    public string Format { get; }

    public ChartImage(byte[] data, string format)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0)
            throw new ArgumentException("image data must be non-empty", nameof(data));
        if (format != FormatSvg && format != FormatPng)
            throw new ArgumentException(
                $"image format must be '{FormatSvg}' or '{FormatPng}', got '{format}'",
                nameof(format));
        Data = data;
        Format = format;
    }

    public override string ToString() => $"[PyExcel chart image: {Format}, {Data.Length} bytes]";
}
