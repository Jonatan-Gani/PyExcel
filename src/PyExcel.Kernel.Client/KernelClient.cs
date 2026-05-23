using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using PyExcel.Bridge;

namespace PyExcel.Kernel.Client;

/// <summary>
/// Typed front-end to <see cref="KernelSupervisor"/>. Builds RUN_REQUEST
/// frames from a <see cref="RunRequest"/>, drives the reply loop until a
/// terminal frame arrives, and either returns a <see cref="RunResult"/>
/// or throws <see cref="KernelException"/>.
///
/// <para>Threading model:</para>
/// <list type="bullet">
///   <item><see cref="Run"/> holds the supervisor's <c>ExchangeSemaphore</c>
///     for the lifetime of the call. Only one Run / Ping / Shutdown runs
///     at a time per supervisor.</item>
///   <item><see cref="Cancel"/> does <em>not</em> take the exchange
///     semaphore — it only acquires the write lock for long enough to push
///     a CANCEL frame, which lets it fire while a Run is parked in a
///     blocking read.</item>
///   <item><see cref="ProgressReceived"/> and <see cref="LogReceived"/>
///     fire synchronously on the Run caller's thread, between frame reads.
///     Handlers must be cheap; long work in a handler stalls the kernel.</item>
/// </list>
///
/// <para>The client does not own the supervisor — disposal of the
/// supervisor is the caller's responsibility. The client itself holds no
/// resources beyond event subscriptions, so it intentionally does not
/// implement <see cref="IDisposable"/>.</para>
/// </summary>
public sealed class KernelClient
{
    private readonly KernelSupervisor _supervisor;

    /// <summary>Fired for each PROGRESS frame received during a Run.</summary>
    public event EventHandler<ProgressReceivedEventArgs>? ProgressReceived;

    /// <summary>Fired for each LOG frame received during a Run.</summary>
    public event EventHandler<LogReceivedEventArgs>? LogReceived;

    public KernelClient(KernelSupervisor supervisor)
    {
        _supervisor = supervisor ?? throw new ArgumentNullException(nameof(supervisor));
    }

    /// <summary>
    /// Synchronously execute one job. Blocks until a terminal frame
    /// (<c>RUN_RESULT</c> or <c>ERROR</c>) arrives, or until the overall
    /// <paramref name="timeoutMs"/> elapses.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="request"/> is null.</exception>
    /// <exception cref="ArgumentException"><see cref="RunRequest.Script"/> is empty.</exception>
    /// <exception cref="TimeoutException">No reply arrived within <paramref name="timeoutMs"/>,
    /// or the kernel was busy with another exchange and didn't free up in time.</exception>
    /// <exception cref="KernelException">The kernel returned an ERROR frame.</exception>
    public RunResult Run(RunRequest request, int timeoutMs = 60_000)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));
        if (string.IsNullOrEmpty(request.Script))
            throw new ArgumentException("RunRequest.Script must be non-empty", nameof(request));
        if (timeoutMs <= 0) throw new ArgumentOutOfRangeException(nameof(timeoutMs));

        var runId = request.RunId ?? Guid.NewGuid().ToString("N");
        var meta = BuildRequestMeta(request, runId);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);

        if (!_supervisor.ExchangeSemaphore.Wait(timeoutMs))
            throw new TimeoutException(
                $"kernel busy with another exchange for {timeoutMs}ms; could not start run {runId}");

        try
        {
            lock (_supervisor.WriteLock)
            {
                _supervisor.Transport.WriteFrame(FrameType.RunRequest, meta, request.Arguments);
            }

            while (true)
            {
                if (DateTime.UtcNow >= deadline)
                    throw new TimeoutException(
                        $"run {runId} did not produce a terminal frame within {timeoutMs}ms");

                Frame frame;
                lock (_supervisor.ReadLock)
                {
                    frame = KernelSupervisor.ReadFrameWithDeadline(_supervisor.Transport, deadline);
                }

                switch (frame.Type)
                {
                    case FrameType.Progress:
                        RaiseProgress(frame, runId);
                        break;
                    case FrameType.Log:
                        RaiseLog(frame, runId);
                        break;
                    case FrameType.RunResult:
                        return BuildRunResult(frame, runId);
                    case FrameType.Error:
                        throw BuildException(frame, runId);
                    default:
                        throw new InvalidOperationException(
                            $"unexpected frame type {frame.Type} during run {runId}");
                }
            }
        }
        finally
        {
            _supervisor.ExchangeSemaphore.Release();
        }
    }

    /// <summary>
    /// Asynchronous wrapper around <see cref="Run"/>. When the
    /// <paramref name="cancellationToken"/> fires, a CANCEL frame is sent
    /// to the kernel; the kernel then replies with an ERROR/<c>Cancelled</c>
    /// frame, which surfaces as <see cref="OperationCanceledException"/>
    /// rather than <see cref="KernelException"/>.
    /// </summary>
    public Task<RunResult> RunAsync(
        RunRequest request,
        CancellationToken cancellationToken = default,
        int timeoutMs = 60_000)
    {
        if (request is null) throw new ArgumentNullException(nameof(request));

        // Pin the run id locally so Cancel can target it and the caller's
        // request stays unmutated.
        var runId = request.RunId ?? Guid.NewGuid().ToString("N");
        var pinned = new RunRequest
        {
            Script = request.Script,
            Function = request.Function,
            Arguments = request.Arguments,
            Kwargs = request.Kwargs,
            RunId = runId,
        };

        return Task.Run(() =>
        {
            using var reg = cancellationToken.Register(() =>
            {
                try { Cancel(runId); } catch { /* best-effort */ }
            });
            try
            {
                return Run(pinned, timeoutMs);
            }
            catch (KernelException ex) when (
                cancellationToken.IsCancellationRequested
                && string.Equals(ex.Code, "Cancelled", StringComparison.Ordinal))
            {
                throw new OperationCanceledException(ex.Message, ex, cancellationToken);
            }
        }, cancellationToken);
    }

    /// <summary>
    /// Fire-and-forget cancel of an in-flight run. Sends a CANCEL frame
    /// carrying <paramref name="runId"/>; the kernel will reply to the
    /// blocked <see cref="Run"/> with an ERROR/<c>Cancelled</c> frame.
    ///
    /// <para>Safe to call from a different thread than <see cref="Run"/> —
    /// it only acquires the write lock, never the exchange semaphore.</para>
    /// </summary>
    public void Cancel(string runId)
    {
        if (string.IsNullOrEmpty(runId))
            throw new ArgumentException("runId required", nameof(runId));

        var meta = new Dictionary<string, object?> { ["run_id"] = runId };
        lock (_supervisor.WriteLock)
        {
            _supervisor.Transport.WriteFrame(FrameType.Cancel, meta);
        }
    }

    // -------------------------------------------------------------------------
    // Request meta + reply parsing
    // -------------------------------------------------------------------------

    private static IReadOnlyDictionary<string, object?> BuildRequestMeta(
        RunRequest req, string runId)
    {
        var meta = new Dictionary<string, object?>
        {
            ["run_id"] = runId,
            ["script"] = req.Script,
            ["function"] = string.IsNullOrEmpty(req.Function) ? "transform" : req.Function,
        };
        if (req.Kwargs is { Count: > 0 })
        {
            // Re-pack as a plain Dictionary so the canonical-JSON encoder
            // walks it as an object via its IDictionary branch.
            // (IReadOnlyDictionary does not implement the non-generic
            // IDictionary interface, and netstandard2.0's Dictionary doesn't
            // have an IEnumerable<KeyValuePair> constructor, so we copy.)
            var copy = new Dictionary<string, object?>(req.Kwargs.Count);
            foreach (var kv in req.Kwargs) copy[kv.Key] = kv.Value;
            meta["kwargs"] = copy;
        }
        return meta;
    }

    private static RunResult BuildRunResult(Frame frame, string runId)
    {
        var echoedRunId = AsString(frame.Meta, "run_id") ?? runId;
        var durationMs = AsInt(frame.Meta, "duration_ms") ?? 0;
        return new RunResult(echoedRunId, durationMs, frame.Payloads);
    }

    private static KernelException BuildException(Frame frame, string runId)
    {
        var echoedRunId = AsString(frame.Meta, "run_id") ?? runId;
        var code = AsString(frame.Meta, "code") ?? "Unknown";
        var pyType = AsString(frame.Meta, "type") ?? code;
        var message = AsString(frame.Meta, "message") ?? "";
        var traceback = AsString(frame.Meta, "traceback") ?? "";
        var durationMs = AsInt(frame.Meta, "duration_ms") ?? 0;
        return new KernelException(echoedRunId, code, pyType, message, traceback, durationMs);
    }

    private void RaiseProgress(Frame frame, string runId)
    {
        var handler = ProgressReceived;
        if (handler is null) return;
        var echoedRunId = AsString(frame.Meta, "run_id") ?? runId;
        var percent = AsDouble(frame.Meta, "percent");
        var message = AsString(frame.Meta, "message");
        handler(this, new ProgressReceivedEventArgs(echoedRunId, percent, message, frame.Meta));
    }

    private void RaiseLog(Frame frame, string runId)
    {
        var handler = LogReceived;
        if (handler is null) return;
        var echoedRunId = AsString(frame.Meta, "run_id") ?? runId;
        var level = AsString(frame.Meta, "level") ?? "info";
        var text = AsString(frame.Meta, "text") ?? "";
        handler(this, new LogReceivedEventArgs(echoedRunId, level, text, frame.Meta));
    }

    // -------------------------------------------------------------------------
    // Meta accessors. Frame meta values are object? boxed primitives — the
    // JSON decoder produces long/double/string/bool/null/dict/list, never
    // the narrower int / float / etc. types — so we coalesce on read.
    // -------------------------------------------------------------------------

    private static string? AsString(IReadOnlyDictionary<string, object?> meta, string key)
        => meta.TryGetValue(key, out var v) && v is string s ? s : null;

    private static int? AsInt(IReadOnlyDictionary<string, object?> meta, string key)
    {
        if (!meta.TryGetValue(key, out var v)) return null;
        return v switch
        {
            long l => (int)l,
            int i => i,
            double d => (int)d,
            _ => null,
        };
    }

    private static double? AsDouble(IReadOnlyDictionary<string, object?> meta, string key)
    {
        if (!meta.TryGetValue(key, out var v)) return null;
        return v switch
        {
            double d => d,
            long l => (double)l,
            int i => (double)i,
            _ => null,
        };
    }
}
