using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using PyExcel.Bridge;
using PyExcel.Kernel.Client;

// Test-only access to the internal classification + decoding helpers in
// this class so PyRunTests can drive them without spinning up a kernel.
[assembly: InternalsVisibleTo("PyExcel.Bridge.Tests")]

namespace PyExcel.Excel;

/// <summary>
/// The marshal-and-dispatch core of the <c>=PY.RUN</c> UDF. Pure-logic
/// boundary so the Excel-DNA wrapper (net48-only, ships in the .xll) can
/// stay one method tall, and unit tests can drive the whole pipeline
/// without spinning up Excel.
///
/// <para>One call:</para>
///
/// <list type="number">
///   <item>Resolve the script path (relative paths are interpreted against
///     the supplied workbook directory; absolute paths pass through).</item>
///   <item>Encode the input — a 2-D <c>object?[,]</c>, 1-D <c>object?[]</c>,
///     or a scalar — into a single Arrow IPC stream via
///     <see cref="ArrowMarshal"/>.</item>
///   <item>Hand the request to <see cref="KernelClient.Run"/> and wait
///     for <c>RUN_RESULT</c> (or <see cref="KernelException"/>).</item>
///   <item>Decode the result back into the shape Excel-DNA can spill —
///     <c>object?[,]</c> for tables, <c>object?[,]</c> sized 1×N or N×1
///     for vectors (depending on Arrow orientation metadata), a boxed
///     scalar otherwise.</item>
/// </list>
///
/// <para>A <c>None</c> return from the user function surfaces as
/// <see cref="EmptyResult"/> — a sentinel the UDF wrapper translates to
/// <c>ExcelDna.Integration.ExcelEmpty.Value</c>. We don't reference
/// ExcelDna here so the netstandard2.0 slice builds cleanly on Linux CI.</para>
/// </summary>
public static class PyRun
{
    /// <summary>Sentinel returned in place of a <c>null</c> when the user
    /// function returned <c>None</c>. The UDF wrapper turns this into
    /// <c>ExcelDna.Integration.ExcelEmpty.Value</c>.</summary>
    public static readonly object EmptyResult = new();

    /// <summary>
    /// Run one job and return the spill-ready result.
    /// </summary>
    /// <param name="script">Path to the user script. Absolute paths are
    /// used as-is; relative paths are resolved against
    /// <paramref name="workbookDirectory"/> when provided, otherwise the
    /// kernel's working directory.</param>
    /// <param name="input">The input range or scalar. Accepts
    /// <c>object?[,]</c>, <c>object?[]</c>, a plain scalar, or
    /// <see langword="null"/> for no-arg invocations.</param>
    /// <param name="kwargs">Optional keyword arguments. Values must be
    /// JSON-serialisable primitives (string, bool, long, double, null) or
    /// nested dicts/lists.</param>
    /// <param name="client">Live kernel client. Phase 4 callers will
    /// almost always pass <see cref="KernelHost.Default"/>.Client.</param>
    /// <param name="workbookDirectory">Optional directory for resolving
    /// relative <paramref name="script"/> paths. Phase 3 wires in the
    /// actual workbook path; Phase 4 may pass <see langword="null"/>.</param>
    /// <param name="function">Function name inside the script. Defaults
    /// to <c>"transform"</c>.</param>
    /// <param name="timeoutMs">Overall budget for the run, including
    /// queueing behind any other in-flight exchange.</param>
    public static object Execute(
        string script,
        object? input,
        IReadOnlyDictionary<string, object?>? kwargs,
        KernelClient client,
        string? workbookDirectory = null,
        string function = "transform",
        int timeoutMs = 60_000)
    {
        if (script is null) throw new ArgumentNullException(nameof(script));
        if (script.Length == 0) throw new ArgumentException("script path must be non-empty", nameof(script));
        if (client is null) throw new ArgumentNullException(nameof(client));

        var scriptPath = ResolveScriptPath(script, workbookDirectory);

        var argBuffer = EncodeInput(input);
        var arguments = argBuffer is null ? Array.Empty<byte[]>() : new[] { argBuffer };

        var result = client.Run(
            new RunRequest
            {
                Script = scriptPath,
                Function = function,
                Arguments = arguments,
                Kwargs = kwargs,
            },
            timeoutMs: timeoutMs);

        return DecodeResult(result);
    }

    // -------------------------------------------------------------------------
    // Input encoding — classify the boxed value, choose the right shape
    // -------------------------------------------------------------------------

    /// <summary>
    /// Encode the caller's input into a single Arrow IPC buffer, or return
    /// <see langword="null"/> if the call has no positional argument.
    /// </summary>
    /// <remarks>
    /// Visibility is <c>internal</c> so the test project can exercise the
    /// classification table without going through a kernel.
    /// </remarks>
    internal static byte[]? EncodeInput(object? input)
    {
        if (input is null) return null;
        // Array covariance: 'object[,]' matches both object[,] and object?[,]
        // at runtime, and ArrowMarshal.EncodeTable's object?[,] parameter
        // accepts the upcast. Avoiding '?' in the pattern keeps us out of
        // an awkward corner of C# 10's pattern-matching grammar.
        if (input is object[,] table) return ArrowMarshal.EncodeTable(table);
        if (input is object[] vector) return ArrowMarshal.EncodeVector(vector);
        return ArrowMarshal.EncodeScalar(input);
    }

    // -------------------------------------------------------------------------
    // Result decoding — translate Arrow shape + orientation into Excel spill
    // -------------------------------------------------------------------------

    /// <summary>
    /// Turn a <see cref="RunResult"/> into a value the UDF can return.
    /// Honours <see cref="ArrowOrientation"/> on vectors: a row vector
    /// becomes a 1×N rectangle (spills across), a column vector becomes
    /// an N×1 rectangle (spills down).
    /// </summary>
    /// <remarks>
    /// Visibility is <c>internal</c> for the same testability reason as
    /// <see cref="EncodeInput"/>.
    /// </remarks>
    internal static object DecodeResult(RunResult result)
    {
        if (result.IsEmpty) return EmptyResult;

        var buffer = result.Payload;
        var (shape, orientation) = ArrowMarshal.PeekShape(buffer);
        var decoded = ArrowMarshal.Decode(buffer);

        if (shape == ArrowShape.Vector && decoded is object?[] vector)
        {
            return SpillVector(vector, orientation ?? ArrowOrientation.Column);
        }
        return decoded ?? EmptyResult;
    }

    /// <summary>
    /// Reshape a 1-D vector into a rectangular array for Excel spill.
    /// Defaults to column-orientation (N×1) when the producer didn't
    /// specify — matches the way Excel-DNA exposes 1-D arrays today.
    /// </summary>
    private static object?[,] SpillVector(object?[] vector, ArrowOrientation orientation)
    {
        if (orientation == ArrowOrientation.Row)
        {
            var row = new object?[1, vector.Length];
            for (var i = 0; i < vector.Length; i++) row[0, i] = vector[i];
            return row;
        }
        var col = new object?[vector.Length, 1];
        for (var i = 0; i < vector.Length; i++) col[i, 0] = vector[i];
        return col;
    }

    // -------------------------------------------------------------------------
    // Path resolution
    // -------------------------------------------------------------------------

    private static string ResolveScriptPath(string script, string? workbookDirectory)
    {
        if (Path.IsPathRooted(script)) return script;
        if (!string.IsNullOrWhiteSpace(workbookDirectory))
            return Path.GetFullPath(Path.Combine(workbookDirectory!, script));
        return script;
    }
}
