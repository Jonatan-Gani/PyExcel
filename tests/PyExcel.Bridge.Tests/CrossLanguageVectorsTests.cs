using System;
using System.Collections.Generic;
using System.Globalization;
using PyExcel.Bridge;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Golden cross-language frame vectors. The Python counterpart lives at
/// <c>tests/kernel/test_cross_language_vectors.py</c>. If both sides pass,
/// C# and Python produce byte-identical encodings for these inputs.
///
/// Vectors deliberately use only types whose JSON encoding is fully
/// specified — null, bool, integer, ASCII string. Float repr depends on
/// the runtime; floats are roundtrip-tested but not byte-pinned.
///
/// If a vector below changes, update both sides in the same commit.
/// </summary>
public class CrossLanguageVectorsTests
{
    public static IEnumerable<object?[]> Vectors() => new[]
    {
        // PING with empty meta, no payloads. Empty dict still serialises to "{}".
        new object?[]
        {
            FrameType.Ping,
            new Dictionary<string, object?>(),
            Array.Empty<byte[]>(),
            "0b00000002020000007b7d00000000",
        },
        // PONG with {"i": 42}.
        new object?[]
        {
            FrameType.Pong,
            new Dictionary<string, object?> { { "i", 42L } },
            Array.Empty<byte[]>(),
            "1100000003080000007b2269223a34327d00000000",
        },
        // ERROR with {"code": "BAD"}.
        new object?[]
        {
            FrameType.Error,
            new Dictionary<string, object?> { { "code", "BAD" } },
            Array.Empty<byte[]>(),
            "17000000050e0000007b22636f6465223a22424144227d00000000",
        },
        // RUN_RESULT with {} meta, single payload 0xdeadbeef.
        new object?[]
        {
            FrameType.RunResult,
            new Dictionary<string, object?>(),
            new[] { new byte[] { 0xde, 0xad, 0xbe, 0xef } },
            "130000000b020000007b7d0100000004000000deadbeef",
        },
        // HELLO with {"kernel":"py", "protocol":2} (alphabetised by sort).
        new object?[]
        {
            FrameType.Hello,
            new Dictionary<string, object?> { { "protocol", 2L }, { "kernel", "py" } },
            Array.Empty<byte[]>(),
            "25000000011c0000007b226b65726e656c223a227079222c2270726f"
            + "746f636f6c223a327d00000000",
        },
        // LOG with {"empty":null, "flag":true, "n":-7}.
        new object?[]
        {
            FrameType.Log,
            new Dictionary<string, object?>
            {
                { "flag", true },
                { "n", -7L },
                { "empty", null },
            },
            Array.Empty<byte[]>(),
            "2a0000000d210000007b22656d707479223a6e756c6c2c22666c6167"
            + "223a747275652c226e223a2d377d00000000",
        },
    };

    [Theory]
    [MemberData(nameof(Vectors))]
    public void EncodeMatchesGolden(
        FrameType ftype,
        Dictionary<string, object?> meta,
        byte[][] payloads,
        string expectedHex)
    {
        var wire = Framing.EncodeFrame(ftype, meta, payloads);
        Assert.Equal(expectedHex, ToHex(wire));
    }

    [Theory]
    [MemberData(nameof(Vectors))]
    public void DecodeMatchesGolden(
        FrameType ftype,
        Dictionary<string, object?> meta,
        byte[][] payloads,
        string expectedHex)
    {
        var wire = FromHex(expectedHex);
        var decoded = Framing.DecodeFrame(wire);

        Assert.Equal(ftype, decoded.Type);
        Assert.Equal(meta.Count, decoded.Meta.Count);
        foreach (var kvp in meta)
        {
            Assert.True(decoded.Meta.ContainsKey(kvp.Key), $"missing key {kvp.Key}");
            Assert.Equal(kvp.Value, decoded.Meta[kvp.Key]);
        }
        Assert.Equal(payloads.Length, decoded.Payloads.Count);
        for (var i = 0; i < payloads.Length; i++)
            Assert.Equal(payloads[i], decoded.Payloads[i]);
    }

    private static string ToHex(byte[] bytes)
    {
        var sb = new System.Text.StringBuilder(bytes.Length * 2);
        foreach (var b in bytes)
            sb.Append(b.ToString("x2", CultureInfo.InvariantCulture));
        return sb.ToString();
    }

    private static byte[] FromHex(string hex)
    {
        if ((hex.Length & 1) != 0)
            throw new ArgumentException("hex string must have even length", nameof(hex));
        var bytes = new byte[hex.Length / 2];
        for (var i = 0; i < bytes.Length; i++)
            bytes[i] = byte.Parse(hex.AsSpan(i * 2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
        return bytes;
    }
}
