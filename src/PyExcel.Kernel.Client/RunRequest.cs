using System;
using System.Collections.Generic;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Inputs to one job. Build one of these, hand it to
/// <see cref="KernelClient.Run"/>, get back a <see cref="RunResult"/> or
/// a <see cref="KernelException"/>.
///
/// <para>All payload bytes are Arrow IPC streams produced by the caller
/// (typically via the <c>PyExcel.Bridge.ArrowIo</c> helper that ships
/// alongside the Python <c>arrow_io.py</c> module). This type itself has
/// no Apache.Arrow dependency — it just shuttles bytes.</para>
/// </summary>
public sealed class RunRequest
{
    /// <summary>
    /// Filesystem path to the user's Python script. The kernel resolves
    /// this on its end, so the path must be valid in the kernel's working
    /// directory (typically the workbook's directory).
    /// </summary>
    public string Script { get; set; } = "";

    /// <summary>
    /// Function name to call inside the script. Defaults to "transform"
    /// to match the PY.RUN convention.
    /// </summary>
    public string Function { get; set; } = "transform";

    /// <summary>
    /// Positional arguments, each one an Arrow IPC stream. The kernel
    /// decodes them in order and passes them to the target function as
    /// <c>fn(*args, **kwargs)</c>.
    /// </summary>
    public IReadOnlyList<byte[]> Arguments { get; set; } = Array.Empty<byte[]>();

    /// <summary>
    /// JSON-serialisable keyword arguments. Values must be primitives
    /// (string, bool, long, double, null) or nested dicts/lists thereof.
    /// </summary>
    public IReadOnlyDictionary<string, object?>? Kwargs { get; set; }

    /// <summary>
    /// Optional caller-provided id for matching <c>Cancel</c> calls to
    /// this run. If null, <see cref="KernelClient.Run"/> generates a Guid
    /// and writes it back to <see cref="RunResult.RunId"/>.
    /// </summary>
    public string? RunId { get; set; }
}
