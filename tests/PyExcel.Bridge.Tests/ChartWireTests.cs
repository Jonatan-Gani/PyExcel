using System;
using System.Collections.Generic;
using System.IO;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using Apache.Arrow.Types;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Wire-level tests for the chart / image Arrow shapes:
/// <see cref="ArrowMarshal.Decode"/> must turn a
/// <c>pyexcel-shape = chart</c> buffer into a typed <see cref="ChartSpec"/>
/// and a <c>pyexcel-shape = image</c> buffer into a
/// <see cref="ChartImage"/>, mirroring what
/// <c>arrow_io.py</c> encodes for Plotly / Matplotlib figure returns.
/// Buffers are built by hand here because the C# encoder intentionally
/// has no chart/image arm (only the kernel produces them); the
/// cross-language conformance check is the e2e pair in
/// <see cref="PyRunTests"/>.
/// </summary>
public class ChartWireTests
{
    // -------------------------------------------------------------------------
    // Typed wire value validation
    // -------------------------------------------------------------------------

    [Fact]
    public void ChartSpec_RejectsNullAndBlank()
    {
        Assert.Throws<ArgumentNullException>(() => new ChartSpec(null!));
        Assert.Throws<ArgumentException>(() => new ChartSpec("   "));
    }

    [Fact]
    public void ChartSpec_EqualityIsOrdinalOnJson()
    {
        Assert.Equal(new ChartSpec("{}"), new ChartSpec("{}"));
        Assert.NotEqual(new ChartSpec("{}"), new ChartSpec("{ }"));
    }

    [Fact]
    public void ChartImage_RejectsBadArguments()
    {
        Assert.Throws<ArgumentNullException>(() => new ChartImage(null!, "png"));
        Assert.Throws<ArgumentException>(() => new ChartImage(System.Array.Empty<byte>(), "png"));
        Assert.Throws<ArgumentException>(() => new ChartImage(new byte[] { 1 }, "jpeg"));
    }

    // -------------------------------------------------------------------------
    // Decode: chart shape
    // -------------------------------------------------------------------------

    [Fact]
    public void Decode_ChartShape_YieldsTypedChartSpec()
    {
        const string json = @"{""version"": 1, ""traces"": []}";
        var buffer = BuildChartBuffer(json);

        var decoded = ArrowMarshal.Decode(buffer);

        var spec = Assert.IsType<ChartSpec>(decoded);
        Assert.Equal(json, spec.Json);
    }

    [Fact]
    public void PeekShape_ChartBuffer_ReportsChart()
    {
        var buffer = BuildChartBuffer(@"{""version"": 1}");
        Assert.Equal(ArrowShape.Chart, ArrowMarshal.PeekShape(buffer).Shape);
    }

    [Fact]
    public void Decode_ChartSpec_ParsesWithChartSpecParser()
    {
        // The full pipeline below the COM boundary: wire bytes → ChartSpec
        // → validated document.
        var buffer = BuildChartBuffer(
            @"{""version"": 1, ""chart_type"": ""bar"", ""traces"": [
                {""id"": 1, ""x"": [""a""], ""y"": [2],
                 ""style"": {""series_type"": ""column""}}]}");

        var spec = (ChartSpec)ArrowMarshal.Decode(buffer)!;
        var doc = ChartSpecParser.Parse(spec.Json);

        Assert.Equal("bar", doc.ChartType);
        Assert.Equal("column", doc.Traces[0].Style.SeriesType);
    }

    // -------------------------------------------------------------------------
    // Decode: image shape
    // -------------------------------------------------------------------------

    [Fact]
    public void Decode_ImageShape_YieldsTypedChartImage()
    {
        var payload = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        var buffer = BuildImageBuffer(payload, "png");

        var decoded = ArrowMarshal.Decode(buffer);

        var image = Assert.IsType<ChartImage>(decoded);
        Assert.Equal(payload, image.Data);
        Assert.Equal("png", image.Format);
    }

    [Fact]
    public void Decode_SvgImage_CarriesFormat()
    {
        var buffer = BuildImageBuffer(new byte[] { (byte)'<' }, "svg");
        var image = (ChartImage)ArrowMarshal.Decode(buffer)!;
        Assert.Equal("svg", image.Format);
    }

    [Fact]
    public void Decode_ImageWithoutFormatMetadata_DefaultsToPng()
    {
        var buffer = BuildImageBuffer(new byte[] { 1, 2, 3 }, format: null);
        var image = (ChartImage)ArrowMarshal.Decode(buffer)!;
        Assert.Equal("png", image.Format);
    }

    [Fact]
    public void PeekShape_ImageBuffer_ReportsImage()
    {
        var buffer = BuildImageBuffer(new byte[] { 1 }, "png");
        Assert.Equal(ArrowShape.Image, ArrowMarshal.PeekShape(buffer).Shape);
    }

    // -------------------------------------------------------------------------
    // Buffer builders (the kernel's encode side, replicated by hand)
    // -------------------------------------------------------------------------

    private static byte[] BuildChartBuffer(string json)
    {
        var builder = new StringArray.Builder();
        builder.Append(json);
        return WriteBuffer(
            new Field("0", StringType.Default, nullable: true),
            builder.Build(),
            shape: "chart");
    }

    private static byte[] BuildImageBuffer(byte[] data, string? format)
    {
        var builder = new BinaryArray.Builder();
        builder.Append(data.AsSpan());

        Dictionary<string, string>? fieldMetadata = null;
        if (format is { })
            fieldMetadata = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["pyexcel-image-format"] = format,
            };

        return WriteBuffer(
            new Field("0", BinaryType.Default, nullable: true, fieldMetadata),
            builder.Build(),
            shape: "image");
    }

    private static byte[] WriteBuffer(Field field, IArrowArray array, string shape)
    {
        var schemaMetadata = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["pyexcel-shape"] = shape,
        };
        var schema = new Schema.Builder().Field(field).Metadata(schemaMetadata).Build();
        var batch = new RecordBatch(schema, new[] { array }, array.Length);

        using var ms = new MemoryStream();
        using (var writer = new ArrowStreamWriter(ms, schema))
        {
            writer.WriteRecordBatch(batch);
            writer.WriteEnd();
        }
        return ms.ToArray();
    }
}
