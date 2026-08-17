using System;
using System.Collections.Generic;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Successful result of one <see cref="KernelClient.Run"/> call.
///
/// <para><see cref="Payloads"/> carries the user function's return value
/// as a single Arrow IPC stream — or is empty when the function returned
/// <c>None</c>. Decoding (Arrow → DataFrame / list / scalar) is done by
/// the caller; this type just shuttles bytes.</para>
///
/// <para>Failures don't come back as a <c>RunResult</c>; they're thrown
/// as <see cref="KernelException"/>.</para>
/// </summary>
public sealed class RunResult
{
    /// <summary>Echo of the request's <c>RunId</c> (auto-generated if the
    /// caller didn't supply one).</summary>
    public string RunId { get; }

    /// <summary>Wall time the kernel reported for the job, in milliseconds.
    /// Excludes the framing round-trip latency on the wire.</summary>
    public int DurationMs { get; }

    /// <summary>
    /// Arrow IPC payloads in wire order. Zero for a <c>None</c> return, one
    /// for a single value, and one per key when the function returned a
    /// dict of named results — in which case <see cref="OutputNames"/> is
    /// the parallel list of keys.
    /// </summary>
    public IReadOnlyList<byte[]> Payloads { get; }

    /// <summary>
    /// Names for each entry of <see cref="Payloads"/>, or
    /// <see langword="null"/> when the function returned a single
    /// anonymous value. Present exactly when the kernel encoded a dict
    /// return, and it is what lets the host route each result to the
    /// output binding of the same name.
    /// </summary>
    public IReadOnlyList<string>? OutputNames { get; }

    /// <summary>True iff the user function returned <c>None</c>.</summary>
    public bool IsEmpty => Payloads.Count == 0;

    /// <summary>Single Arrow payload when the function returned data.
    /// Throws <see cref="InvalidOperationException"/> on an empty result —
    /// gate with <see cref="IsEmpty"/>.</summary>
    public byte[] Payload =>
        Payloads.Count > 0
            ? Payloads[0]
            : throw new InvalidOperationException(
                "RunResult is empty (user function returned None); check IsEmpty first.");

    public RunResult(
        string runId,
        int durationMs,
        IReadOnlyList<byte[]> payloads,
        IReadOnlyList<string>? outputNames = null)
    {
        RunId = runId ?? throw new ArgumentNullException(nameof(runId));
        DurationMs = durationMs;
        Payloads = payloads ?? throw new ArgumentNullException(nameof(payloads));
        OutputNames = outputNames;
    }
}
