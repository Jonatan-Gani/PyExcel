using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using PyExcel.Bridge;
using PyExcel.Kernel.Client;
using PyExcel.State;

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
        int timeoutMs = 60_000,
        RunArchiveContext? archive = null)
    {
        // Single-input is just the degenerate multi-input case: a null
        // input means "no positional argument", a non-null input means
        // "one positional argument".
        var inputs = input is null
            ? Array.Empty<object?>()
            : new[] { input };

        return ExecuteMany(
            script, inputs, kwargs, client, workbookDirectory, function, timeoutMs, archive);
    }

    /// <summary>
    /// Run one job with an arbitrary number of positional arguments and
    /// return the spill-ready result. The kernel matches arguments to the
    /// user function's parameters positionally — so
    /// <c>transform(prices, signals)</c> receives <paramref name="inputs"/>
    /// in list order. This is the entry point the ribbon's
    /// <c>OnRunPython</c> button uses once the Input field is parsed into
    /// multiple range bindings (see
    /// <see cref="PyExcel.State.RibbonRangeParser"/>); the
    /// single-input <see cref="Execute"/> overload is the one the
    /// <c>=PY.RUN</c> UDF calls.
    /// </summary>
    /// <param name="inputs">Ordered positional arguments. Each element may
    /// be <c>object?[,]</c>, <c>object?[]</c>, or a scalar — but not
    /// <see langword="null"/>: a null in the middle of the list would
    /// misalign the remaining positional arguments, so it's rejected. Use
    /// the single-input <see cref="Execute"/> overload (with
    /// <c>input: null</c>) for a no-argument call.</param>
    public static object ExecuteMany(
        string script,
        IReadOnlyList<object?> inputs,
        IReadOnlyDictionary<string, object?>? kwargs,
        KernelClient client,
        string? workbookDirectory = null,
        string function = "transform",
        int timeoutMs = 60_000,
        RunArchiveContext? archive = null,
        IReadOnlyList<RunBinding>? inputBindings = null,
        IReadOnlyList<RunBinding>? outputBindings = null)
    {
        if (script is null) throw new ArgumentNullException(nameof(script));
        if (script.Length == 0) throw new ArgumentException("script path must be non-empty", nameof(script));
        if (inputs is null) throw new ArgumentNullException(nameof(inputs));
        if (client is null) throw new ArgumentNullException(nameof(client));

        var scriptPath = ResolveScriptPath(script, workbookDirectory);
        var arguments = EncodeArguments(inputs);

        var startedAt = DateTimeOffset.UtcNow;
        var stopwatch = Stopwatch.StartNew();
        RunResult? result = null;
        RunArchiveStatus status = RunArchiveStatus.Success;
        KernelErrorRecord? errorRecord = null;
        try
        {
            result = client.Run(
                new RunRequest
                {
                    Script = scriptPath,
                    Function = function,
                    Arguments = arguments,
                    Kwargs = kwargs,
                    Inputs = inputBindings,
                    Outputs = outputBindings,
                },
                timeoutMs: timeoutMs);
        }
        catch (KernelException kex)
        {
            status = string.Equals(kex.Code, "Cancelled", StringComparison.Ordinal)
                ? RunArchiveStatus.Cancelled
                : RunArchiveStatus.Error;
            if (archive is not null)
                errorRecord = BuildKernelRecord(kex, archive.Source, scriptPath, startedAt);
            throw;
        }
        catch (Exception ex)
        {
            status = RunArchiveStatus.Error;
            if (archive is not null)
                errorRecord = BuildHostRecord(ex, archive.Source, scriptPath, startedAt);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            ArchiveBestEffort(archive, startedAt, scriptPath, function, stopwatch.Elapsed,
                arguments, result, errorRecord, status);
        }

        // result is non-null here: the try block either assigned it or
        // re-threw via one of the catch arms.
        return DecodeResult(result!);
    }

    /// <summary>
    /// Async, cancellable counterpart to <see cref="Execute"/>. When
    /// <paramref name="cancellationToken"/> fires, the kernel receives a
    /// <c>CANCEL</c> frame for this run and the task completes with
    /// <see cref="OperationCanceledException"/> instead of returning a
    /// result. Used by the <c>=PY.RUN</c> UDF to translate Excel-DNA's
    /// cancel-on-formula-change into a kernel-side abort.
    /// </summary>
    public static Task<object> ExecuteAsync(
        string script,
        object? input,
        IReadOnlyDictionary<string, object?>? kwargs,
        KernelClient client,
        string? workbookDirectory = null,
        string function = "transform",
        int timeoutMs = 60_000,
        CancellationToken cancellationToken = default,
        RunArchiveContext? archive = null)
    {
        var inputs = input is null
            ? Array.Empty<object?>()
            : new[] { input };

        return ExecuteManyAsync(
            script, inputs, kwargs, client, workbookDirectory, function, timeoutMs,
            cancellationToken, archive);
    }

    /// <summary>
    /// Async, cancellable counterpart to <see cref="ExecuteMany"/>. See
    /// <see cref="ExecuteAsync"/> for the cancellation contract.
    /// </summary>
    public static async Task<object> ExecuteManyAsync(
        string script,
        IReadOnlyList<object?> inputs,
        IReadOnlyDictionary<string, object?>? kwargs,
        KernelClient client,
        string? workbookDirectory = null,
        string function = "transform",
        int timeoutMs = 60_000,
        CancellationToken cancellationToken = default,
        RunArchiveContext? archive = null,
        IReadOnlyList<RunBinding>? inputBindings = null,
        IReadOnlyList<RunBinding>? outputBindings = null)
    {
        if (script is null) throw new ArgumentNullException(nameof(script));
        if (script.Length == 0) throw new ArgumentException("script path must be non-empty", nameof(script));
        if (inputs is null) throw new ArgumentNullException(nameof(inputs));
        if (client is null) throw new ArgumentNullException(nameof(client));

        var scriptPath = ResolveScriptPath(script, workbookDirectory);
        var arguments = EncodeArguments(inputs);

        var startedAt = DateTimeOffset.UtcNow;
        var stopwatch = Stopwatch.StartNew();
        RunResult? result = null;
        RunArchiveStatus status = RunArchiveStatus.Success;
        KernelErrorRecord? errorRecord = null;
        try
        {
            result = await client.RunAsync(
                new RunRequest
                {
                    Script = scriptPath,
                    Function = function,
                    Arguments = arguments,
                    Kwargs = kwargs,
                    Inputs = inputBindings,
                    Outputs = outputBindings,
                },
                cancellationToken: cancellationToken,
                timeoutMs: timeoutMs).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            status = RunArchiveStatus.Cancelled;
            throw;
        }
        catch (KernelException kex)
        {
            status = string.Equals(kex.Code, "Cancelled", StringComparison.Ordinal)
                ? RunArchiveStatus.Cancelled
                : RunArchiveStatus.Error;
            if (archive is not null)
                errorRecord = BuildKernelRecord(kex, archive.Source, scriptPath, startedAt);
            throw;
        }
        catch (Exception ex)
        {
            status = RunArchiveStatus.Error;
            if (archive is not null)
                errorRecord = BuildHostRecord(ex, archive.Source, scriptPath, startedAt);
            throw;
        }
        finally
        {
            stopwatch.Stop();
            ArchiveBestEffort(archive, startedAt, scriptPath, function, stopwatch.Elapsed,
                arguments, result, errorRecord, status);
        }

        // See note in ExecuteMany — result is non-null on the success path.
        return DecodeResult(result!);
    }

    // -------------------------------------------------------------------------
    // Archive helpers — best-effort write to the run archive
    // -------------------------------------------------------------------------

    private static byte[][] EncodeArguments(IReadOnlyList<object?> inputs)
    {
        var arguments = new byte[inputs.Count][];
        for (var i = 0; i < inputs.Count; i++)
        {
            var buffer = EncodeInput(inputs[i]);
            if (buffer is null)
                throw new ArgumentException(
                    $"input at index {i} is null; null positional arguments are not " +
                    $"supported (they would misalign the remaining arguments). " +
                    $"Use Execute / ExecuteAsync(input: null) for a no-argument call.",
                    nameof(inputs));
            arguments[i] = buffer;
        }
        return arguments;
    }

    private static KernelErrorRecord BuildKernelRecord(
        KernelException kex, string source, string scriptPath, DateTimeOffset timestamp)
        => new(
            Timestamp: timestamp,
            Source: source,
            Code: kex.Code,
            PythonType: kex.PythonType,
            Message: kex.Message,
            PythonTraceback: kex.PythonTraceback,
            ScriptPath: scriptPath);

    private static KernelErrorRecord BuildHostRecord(
        Exception ex, string source, string scriptPath, DateTimeOffset timestamp)
        => new(
            Timestamp: timestamp,
            Source: source,
            Code: "HostError",
            PythonType: ex.GetType().Name,
            Message: ex.Message,
            PythonTraceback: ex.ToString(),
            ScriptPath: scriptPath);

    /// <summary>
    /// Persist the run to <paramref name="archive"/>, swallowing any I/O
    /// failures so they can't mask the user-facing result (or
    /// user-facing exception we're already on the way to throwing).
    /// No-op when <paramref name="archive"/> is <see langword="null"/>.
    /// </summary>
    private static void ArchiveBestEffort(
        RunArchiveContext? archive,
        DateTimeOffset startedAt,
        string scriptPath,
        string function,
        TimeSpan duration,
        IReadOnlyList<byte[]> arguments,
        RunResult? result,
        KernelErrorRecord? errorRecord,
        RunArchiveStatus status)
    {
        if (archive is null) return;
        try
        {
            byte[]? output = result is { IsEmpty: false } ? result.Payload : null;
            archive.Archive.Archive(new RunArchiveEntry(
                Timestamp: startedAt,
                WorkbookKey: archive.WorkbookKey,
                ScriptPath: scriptPath,
                Function: function,
                Source: archive.Source,
                Duration: duration,
                Inputs: arguments,
                Output: output,
                Error: errorRecord,
                Status: status));
        }
        catch
        {
            // Best-effort. Losing an archive entry is strictly less bad
            // than masking the result (or the in-flight exception) with
            // an I/O error from this side path.
        }
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
