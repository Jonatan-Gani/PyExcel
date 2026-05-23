using System;
using System.Collections.Generic;
using System.IO;

namespace PyExcel.Bridge;

/// <summary>
/// Wire framing for the PyExcel v2 kernel protocol.
///
/// Frame layout (all integers little-endian, unsigned):
/// <code>
///   +---------+----+----------+-----------+-----------+--------------+
///   | body    | t  | meta_len | meta_json | n_payload | payloads...  |
///   | u32     | u8 | u32      | bytes     | u32       |              |
///   +---------+----+----------+-----------+-----------+--------------+
/// </code>
/// Counterpart to <c>embedded/pyexcel/kernel/framing.py</c> — both sides
/// must encode and decode byte-for-byte identical frames.
///
/// Design rules (mirror the Python side):
/// <list type="bullet">
///   <item><b>Pure stdlib.</b> No third-party JSON dependency — see
///         <see cref="CanonicalJson"/>.</item>
///   <item><b>Bounded.</b> A configurable maximum frame size (default 256
///         MiB) is enforced on both encode and decode so a malformed peer
///         cannot exhaust memory.</item>
///   <item><b>Stream-friendly.</b> <see cref="ReadFrame"/> reads exactly the
///         bytes it needs from the supplied stream.</item>
///   <item><b>Deterministic.</b> Same logical inputs always produce
///         byte-identical frames; sorted JSON keys, no whitespace.</item>
/// </list>
/// </summary>
public static class Framing
{
    /// <summary>On-wire schema version advertised in the Hello frame.</summary>
    public const int ProtocolVersion = 2;

    /// <summary>Default cap on frame body length (256 MiB).</summary>
    public const int DefaultMaxFrameBytes = 256 * 1024 * 1024;

    private const int U8 = 1;
    private const int U32 = 4;

    // -------------------------------------------------------------------------
    // Encode
    // -------------------------------------------------------------------------

    /// <summary>
    /// Serialize a frame to a single byte array ready to write to the
    /// transport.
    /// </summary>
    /// <exception cref="FrameTooLargeException">
    /// Encoded frame body would exceed <paramref name="maxFrameBytes"/>.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="frameType"/> is not a defined <see cref="FrameType"/>,
    /// or <paramref name="meta"/> contains a non-JSON-encodable value.
    /// </exception>
    public static byte[] EncodeFrame(
        FrameType frameType,
        IReadOnlyDictionary<string, object?> meta,
        IReadOnlyList<byte[]>? payloads = null,
        int maxFrameBytes = DefaultMaxFrameBytes)
    {
        if (meta is null) throw new ArgumentNullException(nameof(meta));
        if (!IsKnownFrameType(frameType))
            throw new ArgumentException($"unknown frame type: {(int)frameType}", nameof(frameType));

        payloads ??= Array.Empty<byte[]>();

        var metaBytes = CanonicalJson.Encode(meta);

        long payloadTotal = 0;
        for (var i = 0; i < payloads.Count; i++)
        {
            if (payloads[i] is null)
                throw new ArgumentException($"payload {i} is null", nameof(payloads));
            payloadTotal += payloads[i].Length;
        }

        long bodyLen = (long)U8
                     + U32
                     + metaBytes.Length
                     + U32
                     + (long)payloads.Count * U32
                     + payloadTotal;

        if (bodyLen > maxFrameBytes)
            throw new FrameTooLargeException(
                $"encoded frame body of {bodyLen} bytes exceeds cap of {maxFrameBytes} bytes");

        var totalLen = checked((int)(U32 + bodyLen));
        var output = new byte[totalLen];
        var offset = 0;

        WriteUInt32(output, ref offset, (uint)bodyLen);
        output[offset++] = (byte)frameType;
        WriteUInt32(output, ref offset, (uint)metaBytes.Length);
        Buffer.BlockCopy(metaBytes, 0, output, offset, metaBytes.Length);
        offset += metaBytes.Length;
        WriteUInt32(output, ref offset, (uint)payloads.Count);
        for (var i = 0; i < payloads.Count; i++)
        {
            var p = payloads[i];
            WriteUInt32(output, ref offset, (uint)p.Length);
            Buffer.BlockCopy(p, 0, output, offset, p.Length);
            offset += p.Length;
        }

        return output;
    }

    // -------------------------------------------------------------------------
    // Decode
    // -------------------------------------------------------------------------

    /// <summary>
    /// Read one frame from a stream. Blocks until the full frame is
    /// available or the stream returns EOF mid-frame.
    /// </summary>
    /// <exception cref="TruncatedFrameException">
    /// Stream returned EOF before the frame was complete.
    /// </exception>
    /// <exception cref="FrameTooLargeException">
    /// Peer announced a body larger than <paramref name="maxFrameBytes"/>.
    /// </exception>
    /// <exception cref="MalformedFrameException">
    /// Internal frame lengths are inconsistent or meta is not a JSON object.
    /// </exception>
    public static Frame ReadFrame(Stream input, int maxFrameBytes = DefaultMaxFrameBytes)
    {
        if (input is null) throw new ArgumentNullException(nameof(input));

        var lenBuf = ReadExact(input, U32, "frame length prefix");
        var bodyLen = ReadUInt32(lenBuf, 0);

        if (bodyLen > (uint)maxFrameBytes)
            throw new FrameTooLargeException(
                $"peer announced frame of {bodyLen} bytes; cap is {maxFrameBytes}");
        if (bodyLen < U8 + U32 + U32)
            throw new MalformedFrameException($"frame body length {bodyLen} is below minimum");

        var body = ReadExact(input, (int)bodyLen, "frame body");
        return DecodeBody(body);
    }

    /// <summary>
    /// Decode a frame from a complete in-memory wire buffer (length prefix
    /// included). Convenience for tests and round-trip code paths.
    /// </summary>
    public static Frame DecodeFrame(byte[] wire, int maxFrameBytes = DefaultMaxFrameBytes)
    {
        if (wire is null) throw new ArgumentNullException(nameof(wire));
        using var ms = new MemoryStream(wire, writable: false);
        return ReadFrame(ms, maxFrameBytes);
    }

    private static Frame DecodeBody(byte[] body)
    {
        var bodyLen = body.Length;
        var offset = 0;

        var typeByte = body[offset];
        offset += U8;
        var ftype = (FrameType)typeByte;
        if (!IsKnownFrameType(ftype))
            throw new MalformedFrameException($"unknown frame type byte {typeByte}");

        var metaLen = ReadUInt32(body, offset);
        offset += U32;
        if ((long)offset + metaLen > bodyLen)
            throw new MalformedFrameException(
                $"meta_len {metaLen} would read past end of frame body");

        Dictionary<string, object?> meta;
        if (metaLen == 0)
        {
            meta = new Dictionary<string, object?>(StringComparer.Ordinal);
        }
        else
        {
            var metaBytes = new byte[metaLen];
            Buffer.BlockCopy(body, offset, metaBytes, 0, (int)metaLen);
            offset += (int)metaLen;

            object? decoded;
            try
            {
                decoded = CanonicalJson.Decode(metaBytes);
            }
            catch (FormatException exc)
            {
                throw new MalformedFrameException(
                    $"meta is not valid UTF-8 JSON: {exc.Message}", exc);
            }
            if (decoded is not Dictionary<string, object?> obj)
                throw new MalformedFrameException(
                    $"meta must be a JSON object, got {decoded?.GetType().Name ?? "null"}");
            meta = obj;
        }

        if (offset + U32 > bodyLen)
            throw new MalformedFrameException("frame truncated before payload count");
        var nPayload = ReadUInt32(body, offset);
        offset += U32;

        var payloads = new List<byte[]>((int)nPayload);
        for (uint i = 0; i < nPayload; i++)
        {
            if (offset + U32 > bodyLen)
                throw new MalformedFrameException(
                    $"frame truncated before size of payload {i}");
            var psize = ReadUInt32(body, offset);
            offset += U32;
            if ((long)offset + psize > bodyLen)
                throw new MalformedFrameException(
                    $"payload {i} size {psize} would read past end of frame body");
            var p = new byte[psize];
            if (psize > 0)
                Buffer.BlockCopy(body, offset, p, 0, (int)psize);
            offset += (int)psize;
            payloads.Add(p);
        }

        if (offset != bodyLen)
            throw new MalformedFrameException(
                $"frame has {bodyLen - offset} trailing bytes after parsing");

        return new Frame(ftype, meta, payloads);
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static byte[] ReadExact(Stream input, int n, string what)
    {
        var buf = new byte[n];
        var read = 0;
        while (read < n)
        {
            int r;
            try
            {
                r = input.Read(buf, read, n - read);
            }
            catch (IOException exc)
            {
                throw new TruncatedFrameException(
                    $"I/O failure reading {what}: {exc.Message}", exc);
            }
            if (r == 0)
                throw new TruncatedFrameException(
                    $"EOF reading {what}: expected {n} bytes, got {read}");
            read += r;
        }
        return buf;
    }

    private static void WriteUInt32(byte[] buf, ref int offset, uint value)
    {
        buf[offset] = (byte)value;
        buf[offset + 1] = (byte)(value >> 8);
        buf[offset + 2] = (byte)(value >> 16);
        buf[offset + 3] = (byte)(value >> 24);
        offset += 4;
    }

    private static uint ReadUInt32(byte[] buf, int offset)
    {
        return (uint)(buf[offset]
                    | buf[offset + 1] << 8
                    | buf[offset + 2] << 16
                    | buf[offset + 3] << 24);
    }

    private static bool IsKnownFrameType(FrameType t)
    {
        return t switch
        {
            FrameType.Hello
            or FrameType.Ping
            or FrameType.Pong
            or FrameType.Shutdown
            or FrameType.Error
            or FrameType.RunRequest
            or FrameType.RunResult
            or FrameType.Progress
            or FrameType.Log
            or FrameType.Cancel
            or FrameType.ListJobs => true,
            _ => false,
        };
    }
}
