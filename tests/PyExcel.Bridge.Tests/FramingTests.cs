using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using PyExcel.Bridge;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Mirrors <c>tests/kernel/test_framing.py</c>. Every test here has a
/// Python counterpart — the on-wire format must remain identical.
/// </summary>
public class FramingTests
{
    // -------------------------------------------------------------------------
    // Roundtrip
    // -------------------------------------------------------------------------

    public static IEnumerable<object?[]> RoundtripCases() => new[]
    {
        new object?[]
        {
            FrameType.Hello,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "protocol", (long)Framing.ProtocolVersion }, { "kernel", "abc" } },
            Array.Empty<byte[]>(),
        },
        new object?[] { FrameType.Ping, new Dictionary<string, object?>(), Array.Empty<byte[]>() },
        new object?[] { FrameType.Pong, new Dictionary<string, object?>(), Array.Empty<byte[]>() },
        new object?[]
        {
            FrameType.RunRequest,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "script", "foo.py" }, { "run_id", "r1" } },
            new[] { new byte[] { 0x00, 0x01, 0x02 } },
        },
        new object?[]
        {
            FrameType.RunResult,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "status", "done" }, { "duration", 0.123 } },
            new[] { Encoding.UTF8.GetBytes("table-bytes"), Encoding.UTF8.GetBytes("another-payload-blob") },
        },
        new object?[]
        {
            FrameType.Progress,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "pct", 42L }, { "message", "halfway" } },
            Array.Empty<byte[]>(),
        },
        new object?[]
        {
            FrameType.Log,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "level", "info" }, { "text", "hello" } },
            Array.Empty<byte[]>(),
        },
        new object?[]
        {
            FrameType.Cancel,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "run_id", "r1" } },
            Array.Empty<byte[]>(),
        },
        new object?[]
        {
            FrameType.Shutdown,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "drain", true }, { "timeout_ms", 3000L } },
            Array.Empty<byte[]>(),
        },
        new object?[]
        {
            FrameType.Error,
            new Dictionary<string, object?>(StringComparer.Ordinal)
                { { "code", "PIPE_ACL_FAILURE" }, { "message", "no SID match" } },
            Array.Empty<byte[]>(),
        },
        new object?[]
        {
            FrameType.ListJobs,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "script", "analysis.py" } },
            Array.Empty<byte[]>(),
        },
    };

    [Theory]
    [MemberData(nameof(RoundtripCases))]
    public void Roundtrip(FrameType ftype, Dictionary<string, object?> meta, byte[][] payloads)
    {
        var wire = Framing.EncodeFrame(ftype, meta, payloads);
        var decoded = Framing.DecodeFrame(wire);

        Assert.Equal(ftype, decoded.Type);
        AssertMetaEqual(meta, decoded.Meta);
        Assert.Equal(payloads.Length, decoded.Payloads.Count);
        for (var i = 0; i < payloads.Length; i++)
            Assert.Equal(payloads[i], decoded.Payloads[i]);
    }

    [Fact]
    public void UnicodeMetaRoundtrip()
    {
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal)
        {
            { "name", "Säles" },
            { "emoji", "✓" },
            { "ja", "日本語" },
        };
        var wire = Framing.EncodeFrame(FrameType.Log, meta);
        var decoded = Framing.DecodeFrame(wire);
        AssertMetaEqual(meta, decoded.Meta);
    }

    [Fact]
    public void EmptyMetaRoundtrip()
    {
        var wire = Framing.EncodeFrame(FrameType.Ping, new Dictionary<string, object?>());
        var decoded = Framing.DecodeFrame(wire);
        Assert.Equal(FrameType.Ping, decoded.Type);
        Assert.Empty(decoded.Meta);
        Assert.Empty(decoded.Payloads);
    }

    [Fact]
    public void ZeroLengthPayloadRoundtrip()
    {
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal) { { "k", "v" } };
        var payloads = new[] { Array.Empty<byte>(), Encoding.UTF8.GetBytes("after-empty") };
        var wire = Framing.EncodeFrame(FrameType.RunResult, meta, payloads);
        var decoded = Framing.DecodeFrame(wire);
        Assert.Equal(2, decoded.Payloads.Count);
        Assert.Empty(decoded.Payloads[0]);
        Assert.Equal(Encoding.UTF8.GetBytes("after-empty"), decoded.Payloads[1]);
    }

    [Fact]
    public void LargePayloadRoundtrip()
    {
        var payload = new byte[8 * 1024 * 1024];
        new Random(0).NextBytes(payload);
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal) { { "size", (long)payload.Length } };
        var wire = Framing.EncodeFrame(FrameType.RunResult, meta, new[] { payload });
        var decoded = Framing.DecodeFrame(wire);
        Assert.Single(decoded.Payloads);
        Assert.Equal(payload, decoded.Payloads[0]);
    }

    // -------------------------------------------------------------------------
    // Determinism
    // -------------------------------------------------------------------------

    [Fact]
    public void DeterministicEncoding()
    {
        var a = Framing.EncodeFrame(FrameType.Log,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "b", 2L }, { "a", 1L } });
        var b = Framing.EncodeFrame(FrameType.Log,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "a", 1L }, { "b", 2L } });
        Assert.Equal(a, b);
    }

    [Fact]
    public void EncodingHasNoWhitespacePadding()
    {
        var wire = Framing.EncodeFrame(FrameType.Log,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "a", 1L }, { "b", 2L } });
        var ascii = Encoding.UTF8.GetString(wire);
        Assert.Contains("{\"a\":1,\"b\":2}", ascii);
        Assert.DoesNotContain("\"a\": 1", ascii);
    }

    // -------------------------------------------------------------------------
    // Size cap (encode side)
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeRejectsOversizeFrame()
    {
        const int cap = 1024;
        var tooBig = new byte[cap + 1];
        Assert.Throws<FrameTooLargeException>(() =>
            Framing.EncodeFrame(FrameType.RunResult,
                new Dictionary<string, object?>(),
                new[] { tooBig },
                maxFrameBytes: cap));
    }

    [Fact]
    public void EncodeAcceptsAtCap()
    {
        const int cap = 256;
        // body_len with empty meta is 1 + 4 + 2 + 4 + 4 + N = 15 + N
        var payload = new byte[cap - 15];
        var wire = Framing.EncodeFrame(FrameType.RunResult,
            new Dictionary<string, object?>(),
            new[] { payload },
            maxFrameBytes: cap);
        var decoded = Framing.DecodeFrame(wire, maxFrameBytes: cap);
        Assert.Single(decoded.Payloads);
        Assert.Equal(payload.Length, decoded.Payloads[0].Length);
    }

    // -------------------------------------------------------------------------
    // Size cap (decode side)
    // -------------------------------------------------------------------------

    [Fact]
    public void DecodeRejectsOversizeAnnouncement()
    {
        const int cap = 1024;
        uint announcedBody = 10 * 1024 * 1024;
        var bogus = new byte[4];
        bogus[0] = (byte)announcedBody;
        bogus[1] = (byte)(announcedBody >> 8);
        bogus[2] = (byte)(announcedBody >> 16);
        bogus[3] = (byte)(announcedBody >> 24);
        Assert.Throws<FrameTooLargeException>(() =>
            Framing.DecodeFrame(bogus, maxFrameBytes: cap));
    }

    // -------------------------------------------------------------------------
    // Truncation handling
    // -------------------------------------------------------------------------

    [Fact]
    public void TruncatedBeforeLengthPrefix()
    {
        Assert.Throws<TruncatedFrameException>(() => Framing.DecodeFrame(Array.Empty<byte>()));
    }

    [Fact]
    public void TruncatedLengthPrefix()
    {
        Assert.Throws<TruncatedFrameException>(() => Framing.DecodeFrame(new byte[] { 0x01, 0x02 }));
    }

    [Fact]
    public void TruncatedBody()
    {
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal) { { "k", "v" } };
        var wire = Framing.EncodeFrame(FrameType.RunResult, meta, new[] { Encoding.UTF8.GetBytes("abc") });
        var truncated = wire.Take(wire.Length - 1).ToArray();
        Assert.Throws<TruncatedFrameException>(() => Framing.DecodeFrame(truncated));
    }

    // -------------------------------------------------------------------------
    // Malformed-frame detection
    // -------------------------------------------------------------------------

    [Fact]
    public void MalformedUnknownFrameType()
    {
        var body = new List<byte> { 99 };
        AddU32(body, 0);
        AddU32(body, 0);
        var wire = WireWith(body);
        var exc = Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
        Assert.Contains("unknown frame type", exc.Message);
    }

    [Fact]
    public void MalformedMetaLenOverrunsBody()
    {
        var body = new List<byte> { (byte)FrameType.Ping };
        AddU32(body, 100);
        AddU32(body, 0);
        var wire = WireWith(body);
        Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
    }

    [Fact]
    public void MalformedMetaNotJson()
    {
        var badMeta = new byte[] { 0xff, 0xfe, (byte)' ', (byte)'n', (byte)'o', (byte)'t' };
        var body = new List<byte> { (byte)FrameType.Log };
        AddU32(body, (uint)badMeta.Length);
        body.AddRange(badMeta);
        AddU32(body, 0);
        var wire = WireWith(body);
        var exc = Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
        Assert.Contains("meta is not valid UTF-8 JSON", exc.Message);
    }

    [Fact]
    public void MalformedMetaNotObject()
    {
        var arrBytes = Encoding.UTF8.GetBytes("[1,2,3]");
        var body = new List<byte> { (byte)FrameType.Log };
        AddU32(body, (uint)arrBytes.Length);
        body.AddRange(arrBytes);
        AddU32(body, 0);
        var wire = WireWith(body);
        var exc = Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
        Assert.Contains("meta must be a JSON object", exc.Message);
    }

    [Fact]
    public void MalformedPayloadSizeOverruns()
    {
        var metaBytes = Encoding.UTF8.GetBytes("{}");
        var body = new List<byte> { (byte)FrameType.RunResult };
        AddU32(body, (uint)metaBytes.Length);
        body.AddRange(metaBytes);
        AddU32(body, 1);
        AddU32(body, 1000);
        body.Add(0x00);
        var wire = WireWith(body);
        Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
    }

    [Fact]
    public void MalformedBodyMinimum()
    {
        var wire = new byte[] { 0x03, 0x00, 0x00, 0x00, 0x01, 0x02, 0x03 };
        var exc = Assert.Throws<MalformedFrameException>(() => Framing.DecodeFrame(wire));
        Assert.Contains("below minimum", exc.Message);
    }

    // -------------------------------------------------------------------------
    // Encoding-side type checks
    // -------------------------------------------------------------------------

    [Fact]
    public void EncodeRejectsNullPayload()
    {
        var meta = new Dictionary<string, object?>();
        var payloads = new byte[1][];
        payloads[0] = null!;
        Assert.Throws<ArgumentException>(() =>
            Framing.EncodeFrame(FrameType.RunResult, meta, payloads));
    }

    [Fact]
    public void EncodeRejectsUnknownFrameTypeInt()
    {
        var meta = new Dictionary<string, object?>();
        var exc = Assert.Throws<ArgumentException>(() =>
            Framing.EncodeFrame((FrameType)99, meta));
        Assert.Contains("unknown frame type", exc.Message);
    }

    [Fact]
    public void EncodeRejectsNonJsonMeta()
    {
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal) { { "obj", new object() } };
        Assert.Throws<ArgumentException>(() => Framing.EncodeFrame(FrameType.Log, meta));
    }

    // -------------------------------------------------------------------------
    // Streaming behaviour
    // -------------------------------------------------------------------------

    [Fact]
    public void BackToBackFramesStream()
    {
        var a = Framing.EncodeFrame(FrameType.Ping,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "i", 1L } });
        var b = Framing.EncodeFrame(FrameType.Log,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "text", "two" } },
            new[] { Encoding.UTF8.GetBytes("payload") });
        var c = Framing.EncodeFrame(FrameType.Pong,
            new Dictionary<string, object?>(StringComparer.Ordinal) { { "i", 1L } });

        using var ms = new MemoryStream(a.Concat(b).Concat(c).ToArray());
        var f1 = Framing.ReadFrame(ms);
        var f2 = Framing.ReadFrame(ms);
        var f3 = Framing.ReadFrame(ms);

        Assert.Equal(FrameType.Ping, f1.Type);
        Assert.Equal(1L, f1.Meta["i"]);
        Assert.Equal(FrameType.Log, f2.Type);
        Assert.Single(f2.Payloads);
        Assert.Equal(Encoding.UTF8.GetBytes("payload"), f2.Payloads[0]);
        Assert.Equal(FrameType.Pong, f3.Type);
    }

    // -------------------------------------------------------------------------
    // Protocol version
    // -------------------------------------------------------------------------

    [Fact]
    public void ProtocolVersionMatchesPython()
    {
        Assert.True(Framing.ProtocolVersion >= 2);
    }

    // -------------------------------------------------------------------------
    // Property-style fuzz
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)] [InlineData(4)]
    [InlineData(5)] [InlineData(6)] [InlineData(7)] [InlineData(8)] [InlineData(9)]
    [InlineData(10)] [InlineData(11)] [InlineData(12)] [InlineData(13)] [InlineData(14)]
    [InlineData(15)] [InlineData(16)] [InlineData(17)] [InlineData(18)] [InlineData(19)]
    public void FuzzRoundtrip(int seed)
    {
        var rng = new Random(seed);
        var allTypes = (FrameType[])Enum.GetValues(typeof(FrameType));
        var ftype = allTypes[rng.Next(allTypes.Length)];
        var meta = RandomMeta(rng);
        var payloads = RandomPayloads(rng);

        var wire = Framing.EncodeFrame(ftype, meta, payloads);
        var decoded = Framing.DecodeFrame(wire);

        Assert.Equal(ftype, decoded.Type);
        AssertMetaEqual(meta, decoded.Meta);
        Assert.Equal(payloads.Length, decoded.Payloads.Count);
        for (var i = 0; i < payloads.Length; i++)
            Assert.Equal(payloads[i], decoded.Payloads[i]);
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static void AssertMetaEqual(
        IReadOnlyDictionary<string, object?> expected,
        IReadOnlyDictionary<string, object?> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        foreach (var kvp in expected)
        {
            Assert.True(actual.ContainsKey(kvp.Key), $"missing key {kvp.Key}");
            AssertValueEqual(kvp.Value, actual[kvp.Key]);
        }
    }

    private static void AssertValueEqual(object? expected, object? actual)
    {
        if (expected is null) { Assert.Null(actual); return; }
        switch (expected)
        {
            case int i:
                Assert.Equal((long)i, actual);
                return;
            case long l:
                Assert.Equal(l, actual);
                return;
            case double d:
                Assert.Equal(d, Assert.IsType<double>(actual), 9);
                return;
            case List<object?> exList:
                var acList = Assert.IsType<List<object?>>(actual);
                Assert.Equal(exList.Count, acList.Count);
                for (var i = 0; i < exList.Count; i++)
                    AssertValueEqual(exList[i], acList[i]);
                return;
            default:
                Assert.Equal(expected, actual);
                return;
        }
    }

    private static void AddU32(List<byte> buf, uint v)
    {
        buf.Add((byte)v);
        buf.Add((byte)(v >> 8));
        buf.Add((byte)(v >> 16));
        buf.Add((byte)(v >> 24));
    }

    private static byte[] WireWith(List<byte> body)
    {
        var wire = new List<byte>();
        AddU32(wire, (uint)body.Count);
        wire.AddRange(body);
        return wire.ToArray();
    }

    private static Dictionary<string, object?> RandomMeta(Random rng)
    {
        var keys = new[] { "a", "b", "c", "name", "id", "value", "kv", "nested" };
        var n = rng.Next(0, 7);
        var meta = new Dictionary<string, object?>(StringComparer.Ordinal);
        for (var i = 0; i < n; i++)
        {
            var k = keys[rng.Next(keys.Length)];
            object? value = rng.Next(6) switch
            {
                0 => (long)rng.Next(int.MinValue, int.MaxValue),
                1 => rng.NextDouble() * 1e6,
                2 => RandomString(rng),
                3 => rng.Next(2) == 0,
                4 => null,
                _ => RandomList(rng),
            };
            meta[k] = value;
        }
        return meta;
    }

    private static string RandomString(Random rng)
    {
        const string alphabet = "abcdef日本✓";
        var len = rng.Next(0, 17);
        var sb = new StringBuilder(len);
        for (var i = 0; i < len; i++)
            sb.Append(alphabet[rng.Next(alphabet.Length)]);
        return sb.ToString();
    }

    private static List<object?> RandomList(Random rng)
    {
        var len = rng.Next(0, 5);
        var list = new List<object?>(len);
        for (var i = 0; i < len; i++)
            list.Add((long)rng.Next(0, 100));
        return list;
    }

    private static byte[][] RandomPayloads(Random rng)
    {
        var n = rng.Next(0, 5);
        var result = new byte[n][];
        for (var i = 0; i < n; i++)
        {
            var len = rng.Next(0, 257);
            var p = new byte[len];
            rng.NextBytes(p);
            result[i] = p;
        }
        return result;
    }
}
