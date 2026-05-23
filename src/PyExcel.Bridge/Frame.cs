using System;
using System.Collections.Generic;

namespace PyExcel.Bridge;

/// <summary>
/// One decoded wire frame.
///
/// <see cref="Meta"/> is the parsed JSON object — values are
/// <see cref="long"/>, <see cref="double"/>, <see cref="string"/>,
/// <see cref="bool"/>, <c>null</c>, <see cref="Dictionary{TKey,TValue}"/>,
/// or <see cref="List{T}"/>. <see cref="Payloads"/> holds the raw binary
/// blobs in wire order (typically a single Arrow IPC stream per run).
/// </summary>
public sealed class Frame
{
    public FrameType Type { get; }
    public IReadOnlyDictionary<string, object?> Meta { get; }
    public IReadOnlyList<byte[]> Payloads { get; }

    public Frame(
        FrameType type,
        IReadOnlyDictionary<string, object?> meta,
        IReadOnlyList<byte[]>? payloads = null)
    {
        Type = type;
        Meta = meta ?? throw new ArgumentNullException(nameof(meta));
        Payloads = payloads ?? Array.Empty<byte[]>();
    }
}
