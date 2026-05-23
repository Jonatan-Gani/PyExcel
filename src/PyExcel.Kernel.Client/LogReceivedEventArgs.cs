using System;
using System.Collections.Generic;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Fired for each <c>LOG</c> frame the kernel emits during a
/// <see cref="KernelClient.Run"/>. Used to surface user <c>print()</c>
/// / <c>logging</c> output that would otherwise be invisible because the
/// kernel's stdout is captured by the host.
/// </summary>
public sealed class LogReceivedEventArgs : EventArgs
{
    /// <summary>Run id this log line belongs to.</summary>
    public string RunId { get; }

    /// <summary>Severity level — one of <c>"debug"</c>, <c>"info"</c>,
    /// <c>"warning"</c>, <c>"error"</c>. Defaults to <c>"info"</c> when the
    /// producer omits it.</summary>
    public string Level { get; }

    /// <summary>The log line itself.</summary>
    public string Text { get; }

    /// <summary>Full frame meta, for producers that include structured fields
    /// beyond level/text.</summary>
    public IReadOnlyDictionary<string, object?> Meta { get; }

    public LogReceivedEventArgs(
        string runId,
        string level,
        string text,
        IReadOnlyDictionary<string, object?> meta)
    {
        RunId = runId ?? "";
        Level = level ?? "info";
        Text = text ?? "";
        Meta = meta ?? throw new ArgumentNullException(nameof(meta));
    }
}
