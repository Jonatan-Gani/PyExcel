using System;
using System.IO;
using System.Text;

namespace PyExcel.State;

/// <summary>
/// One captured error from a kernel run — either a <c>KernelException</c>
/// (kernel-side: bad script, raised <c>ValueError</c>, …) or a host-side
/// exception during marshalling / range I/O. Immutable; safe to pass
/// across threads without copying.
///
/// <para>The ribbon's "Copy Last Error" / "Show Last Error" buttons use
/// <see cref="FormatForClipboard"/> to render this into a multi-line
/// string the user can paste into a bug report or share with a
/// colleague.</para>
/// </summary>
/// <param name="Timestamp">When the error was captured (UTC).</param>
/// <param name="Source">Human-readable origin label, e.g. <c>"PY.RUN"</c>
/// or <c>"Run Python button"</c>.</param>
/// <param name="Code">Kernel error code (<c>"Exception"</c>,
/// <c>"ModuleNotFound"</c>, <c>"Cancelled"</c>, …) or <c>"HostError"</c>
/// for non-kernel failures.</param>
/// <param name="PythonType">The Python exception type name
/// (<c>"ValueError"</c>, <c>"KeyError"</c>, …) when known; same as
/// <paramref name="Code"/> for host-side faults.</param>
/// <param name="Message">The exception's message.</param>
/// <param name="PythonTraceback">Full Python traceback if available;
/// empty for host-side faults.</param>
/// <param name="ScriptPath">The script path the run was targeting,
/// if known. Used to route the error to the right per-script log
/// later (Phase 8).</param>
public sealed record KernelErrorRecord(
    DateTimeOffset Timestamp,
    string Source,
    string Code,
    string PythonType,
    string Message,
    string PythonTraceback,
    string? ScriptPath)
{
    /// <summary>
    /// Render the record into a paste-friendly block. Stable layout —
    /// the ribbon button copies this verbatim, so a future change here
    /// is a visible change to users sharing tracebacks.
    /// </summary>
    public string FormatForClipboard()
    {
        var sb = new StringBuilder();
        sb.AppendLine($"[{Timestamp:yyyy-MM-dd HH:mm:ss}] {Source}");
        sb.AppendLine($"  code:    {Code}");
        if (!string.Equals(PythonType, Code, StringComparison.Ordinal))
            sb.AppendLine($"  type:    {PythonType}");
        sb.AppendLine($"  message: {Message}");
        if (!string.IsNullOrEmpty(ScriptPath))
            sb.AppendLine($"  script:  {ScriptPath}");
        if (!string.IsNullOrEmpty(PythonTraceback))
        {
            sb.AppendLine("  traceback:");
            // Indent the traceback so it visually attaches to the header
            // instead of starting at column 0 like a separate message.
            using var reader = new StringReader(PythonTraceback);
            string? line;
            while ((line = reader.ReadLine()) is not null)
                sb.AppendLine($"    {line}");
        }
        return sb.ToString();
    }
}
