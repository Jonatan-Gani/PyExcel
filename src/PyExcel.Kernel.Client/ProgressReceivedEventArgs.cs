using System;
using System.Collections.Generic;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Fired for each <c>PROGRESS</c> frame the kernel emits during a
/// <see cref="KernelClient.Run"/>. Handlers run on the calling thread of
/// <c>Run</c>, between frame reads — keep them cheap so they don't hold
/// up the kernel.
/// </summary>
public sealed class ProgressReceivedEventArgs : EventArgs
{
    /// <summary>Run id this progress update belongs to.</summary>
    public string RunId { get; }

    /// <summary>Progress as a 0–100 percentage. <c>null</c> for indeterminate
    /// updates that only carry a status message.</summary>
    public double? Percent { get; }

    /// <summary>Optional human-readable status message.</summary>
    public string? Message { get; }

    /// <summary>The raw frame meta, in case the producer included additional
    /// fields beyond <c>percent</c> / <c>message</c>.</summary>
    public IReadOnlyDictionary<string, object?> Meta { get; }

    public ProgressReceivedEventArgs(
        string runId,
        double? percent,
        string? message,
        IReadOnlyDictionary<string, object?> meta)
    {
        RunId = runId ?? "";
        Percent = percent;
        Message = message;
        Meta = meta ?? throw new ArgumentNullException(nameof(meta));
    }
}
