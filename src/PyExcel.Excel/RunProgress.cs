using System;
using System.Threading;

namespace PyExcel.Excel;

/// <summary>
/// The progress surface a long-running script reports against — the
/// abstraction between <see cref="RangeRunner"/> (which subscribes to the
/// kernel's <c>PROGRESS</c> frames) and the modeless WinForms progress
/// dialog (which renders them and offers Cancel). Cross-platform so
/// <c>RangeRunner</c> takes no WinForms dependency; the form in
/// <c>PyExcel.Forms</c> implements it.
/// </summary>
public interface IRunProgressSink
{
    /// <summary>Report progress. <paramref name="percent"/> is null for an
    /// indeterminate step.</summary>
    void Report(double? percent, string message);

    /// <summary>The run finished (success or failure) — tear the UI down.</summary>
    void Complete();

    /// <summary>Cancelled when the user clicks the dialog's Cancel button;
    /// threaded into the async run so the kernel gets a <c>CANCEL</c>
    /// frame.</summary>
    CancellationToken CancellationToken { get; }
}

/// <summary>
/// Pure formatting helpers for the progress dialog — clamping a reported
/// percent into the bar's range and rendering the status line. Kept apart
/// from the WinForms form so the rules are unit-tested on Linux CI.
/// </summary>
public static class ProgressModel
{
    /// <summary>Clamp a reported percent to the 0–100 integer range the
    /// progress bar accepts (kernels may over- or under-shoot).</summary>
    public static int ClampPercent(double percent)
    {
        if (double.IsNaN(percent)) return 0;
        if (percent < 0) return 0;
        if (percent > 100) return 100;
        return (int)Math.Round(percent, MidpointRounding.AwayFromZero);
    }

    /// <summary>Render the status line: "<c>42% — message</c>", or just the
    /// message for an indeterminate step, or a default when both are
    /// absent.</summary>
    public static string FormatLine(double? percent, string? message)
    {
        var msg = (message ?? string.Empty).Trim();
        if (percent is null)
            return msg.Length == 0 ? "Working…" : msg;

        var pct = ClampPercent(percent.Value);
        return msg.Length == 0 ? $"{pct}%" : $"{pct}% — {msg}";
    }
}
