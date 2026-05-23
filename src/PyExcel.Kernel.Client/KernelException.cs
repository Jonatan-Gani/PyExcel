using System;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Thrown by <see cref="KernelClient.Run"/> when the kernel returns an
/// ERROR frame instead of RUN_RESULT.
///
/// <para>The kernel's typed error codes survive the wire crossing in
/// <see cref="Code"/> so callers can branch on them without parsing the
/// message string. See <c>worker.py</c> for the canonical list:
/// <c>BadRequest</c>, <c>ModuleNotFound</c>, <c>ModuleExecError</c>,
/// <c>FunctionNotFound</c>, <c>FunctionNotCallable</c>, <c>BadInput</c>,
/// <c>BadReturnType</c>, <c>Exception</c>, <c>Cancelled</c>.</para>
/// </summary>
public sealed class KernelException : Exception
{
    /// <summary>Echo of the request's run id (may be empty if the kernel
    /// rejected the request before parsing meta).</summary>
    public string RunId { get; }

    /// <summary>Stable error code from <c>worker.py</c>'s
    /// <c>JobError.code</c> taxonomy.</summary>
    public string Code { get; }

    /// <summary>Python exception class name (e.g. <c>"ValueError"</c>) when
    /// <see cref="Code"/> is <c>"Exception"</c>; the code itself otherwise.</summary>
    public string PythonType { get; }

    /// <summary>Formatted Python traceback. Empty when the failure was a
    /// kernel-side validation (<c>BadRequest</c>) rather than a raise inside
    /// user code.</summary>
    public string PythonTraceback { get; }

    /// <summary>Wall time the kernel spent before failing, in milliseconds.</summary>
    public int DurationMs { get; }

    public KernelException(
        string runId,
        string code,
        string pythonType,
        string message,
        string pythonTraceback,
        int durationMs)
        : base(BuildMessage(code, pythonType, message))
    {
        RunId = runId ?? "";
        Code = code ?? "Unknown";
        PythonType = pythonType ?? "";
        PythonTraceback = pythonTraceback ?? "";
        DurationMs = durationMs;
    }

    private static string BuildMessage(string code, string pythonType, string message)
    {
        if (string.IsNullOrEmpty(pythonType) || pythonType == code)
            return $"[{code}] {message}";
        return $"[{code}] {pythonType}: {message}";
    }
}
