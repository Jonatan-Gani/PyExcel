#if NETFRAMEWORK
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.Kernel.Client;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// The <c>=PY.RUN</c> worksheet function — Excel-DNA surface on top of
/// <see cref="PyRun"/>.
///
/// <para>The flow on each cell calc:</para>
///
/// <list type="number">
///   <item>Excel calls <see cref="Run"/> on the calc thread with a script
///     path, one input argument (cell / range / array / scalar), and an
///     optional function name.</item>
///   <item>We hand the work to <see cref="ExcelAsyncUtil.Observe"/> via a
///     small <see cref="IExcelObservable"/>. The first call kicks off a
///     background task and returns <c>#N/A</c> immediately so Excel's UI
///     doesn't freeze (SAFE-1). When the task completes, the cell
///     auto-refreshes with the real result. Subsequent calls with the same
///     parameter set reuse the cached result.</item>
///   <item>On the background task, <see cref="PyRun.ExecuteAsync"/>
///     translates Excel-DNA's sentinel argument types into plain .NET,
///     dispatches against <see cref="KernelHost.Default"/>, and converts
///     the result back into something Excel can spill.</item>
///   <item>If Excel-DNA disposes the subscription (the user changed the
///     formula, the workbook closed, or the cell was deleted), the
///     observable's <see cref="IDisposable.Dispose"/> cancels the
///     <see cref="CancellationTokenSource"/> driving the run. The token
///     registration inside <see cref="KernelClient.RunAsync"/> then pushes
///     a <c>CANCEL</c> frame to the kernel; the user's <c>transform()</c>
///     observes <c>pyexcel.kernel.is_cancelled()</c> and returns early.
///     Either way (cooperative abort or natural completion), the run is
///     no longer holding the supervisor's exchange semaphore, so the next
///     call can proceed immediately.</item>
/// </list>
///
/// <para>Failures surface as <see cref="ExcelError.ExcelErrorValue"/>
/// (<c>#VALUE!</c> in the cell) with the full Python traceback logged
/// to <see cref="Trace"/> and Excel-DNA's <see cref="LogDisplay"/>. A
/// richer error UI — a per-script error pane or a hover tooltip — is a
/// follow-up.</para>
/// </summary>
public static class PyRunFunction
{
    /// <summary>
    /// The <c>=PY.RUN</c> worksheet function. Method name is <c>Run</c>
    /// to avoid colliding with <see cref="PyRun"/> in the same namespace;
    /// Excel sees the function as <c>PY.RUN</c> via the attribute.
    /// </summary>
    [ExcelFunction(
        Name = "PY.RUN",
        Description = "Run a Python transform() function on a range and return its result.",
        Category = "PyExcel",
        IsThreadSafe = false)]
    public static object Run(
        [ExcelArgument(
            Name = "script",
            Description = "Path to the Python script (.py)")]
        string script,
        [ExcelArgument(
            Name = "input",
            Description = "Cell, range, or array passed as the first positional argument",
            AllowReference = false)]
        object input,
        [ExcelArgument(
            Name = "function",
            Description = "Function name in the script (default: transform)")]
        object function)
    {
        // ExcelAsyncUtil.Observe is the cancellable cousin of
        // ExcelAsyncUtil.Run. The parameter tuple keys Excel-DNA's
        // internal cache; identical inputs short-circuit to the cached
        // result rather than re-spawning a job. When Excel-DNA discards
        // a cached observable (user typed a new formula, workbook closed,
        // cell deleted), it disposes the IDisposable returned by
        // Subscribe — that's our hook to abort the kernel run.
        // Positional args here because the ExcelDna 1.8.0 signature uses
        // different parameter names than I'd guessed, and positions are
        // the stable contract.
        return ExcelAsyncUtil.Observe(
            "PY.RUN",
            new object[] { script, input, function },
            () => new PyRunObservable(script, input, function));
    }

    // -------------------------------------------------------------------------
    // Observable that drives one run and cancels on dispose
    // -------------------------------------------------------------------------

    private sealed class PyRunObservable : IExcelObservable
    {
        private readonly string _script;
        private readonly object _input;
        private readonly object _function;

        public PyRunObservable(string script, object input, object function)
        {
            _script = script;
            _input = input;
            _function = function;
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            var cts = new CancellationTokenSource();
            // Kick off the run on a background task. PyRun.ExecuteAsync
            // wires the token into KernelClient.RunAsync, which on cancel
            // pushes a CANCEL frame to the kernel.
            //
            // Task is fire-and-forget by design — all exceptions are
            // caught inside the lambda so the runtime never sees an
            // unobserved task fault. We don't store the task: the only
            // handle we need into it is the CTS, which Dispose owns.
            _ = Task.Run(() => RunCoreAsync(observer, cts.Token));
            return new CancelOnDispose(cts);
        }

        private async Task RunCoreAsync(IExcelObserver observer, CancellationToken token)
        {
            try
            {
                var functionName = ResolveFunctionName(_function);
                var inp = FromExcelArgument(_input);
                var archiveContext = BuildArchiveContext();

                var result = await PyRun.ExecuteAsync(
                    script: _script,
                    input: inp,
                    kwargs: null,
                    client: KernelHost.Default.Client,
                    function: functionName,
                    cancellationToken: token,
                    archive: archiveContext).ConfigureAwait(false);

                if (token.IsCancellationRequested) return; // observer is gone
                observer.OnNext(ToExcelOutput(result));
                observer.OnCompleted();
            }
            catch (OperationCanceledException) when (token.IsCancellationRequested)
            {
                // Excel-DNA cancelled us (formula change, workbook close);
                // the observer is no longer interested. Drop silently.
            }
            catch (KernelException kex)
            {
                // The kernel-side traceback is the diagnostic users actually
                // need. Send it to:
                //   * Trace — visible in DebugView when developing.
                //   * LogDisplay — Excel-DNA's built-in error window, which
                //     users can open from the add-in (or which pops up
                //     automatically on the first message in some configs).
                //   * ErrorService — backs the ribbon's "Show / Copy Last
                //     Error" buttons so the user can recover the traceback
                //     without hunting through Excel-DNA's log window.
                // The cell itself still gets #VALUE! so spreadsheet formulas
                // like ISERROR() see it as a failure rather than as data.
                var record = new KernelErrorRecord(
                    Timestamp: DateTimeOffset.UtcNow,
                    Source: "PY.RUN",
                    Code: kex.Code,
                    PythonType: kex.PythonType,
                    Message: kex.Message,
                    PythonTraceback: kex.PythonTraceback,
                    ScriptPath: _script);
                RecordError(record);
                var msg = record.FormatForClipboard();
                Trace.WriteLine(msg);
                LogDisplay.WriteLine(msg);
                if (token.IsCancellationRequested) return;
                observer.OnNext(ExcelError.ExcelErrorValue);
                observer.OnCompleted();
            }
            catch (Exception ex)
            {
                var record = new KernelErrorRecord(
                    Timestamp: DateTimeOffset.UtcNow,
                    Source: "PY.RUN",
                    Code: "HostError",
                    PythonType: ex.GetType().Name,
                    Message: ex.Message,
                    PythonTraceback: ex.ToString(),
                    ScriptPath: _script);
                RecordError(record);
                var msg = record.FormatForClipboard();
                Trace.WriteLine(msg);
                LogDisplay.WriteLine(msg);
                if (token.IsCancellationRequested) return;
                observer.OnNext(ExcelError.ExcelErrorValue);
                observer.OnCompleted();
            }
        }

        /// <summary>Push the error into the per-workbook last-error slot
        /// (or the global slot if no workbook is active). Best-effort —
        /// a failure here must not eat the user-facing #VALUE!.</summary>
        private static void RecordError(KernelErrorRecord record)
        {
            try
            {
                var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
                PyExcelServices.Errors.Record(key, record);
            }
            catch
            {
                // Best-effort. The user already sees #VALUE! in the cell
                // and the message in LogDisplay; losing the ribbon's
                // copy-button content is a worse experience than crashing
                // the ribbon would be.
            }
        }

        /// <summary>Build the archive context for this UDF call. Returns
        /// <see langword="null"/> if the active <see cref="RunArchive"/>
        /// can't be obtained — archiving is opt-in diagnostic data, not
        /// load-bearing, so a missing one degrades gracefully.</summary>
        private static RunArchiveContext? BuildArchiveContext()
        {
            try
            {
                var archive = PyExcelServices.RunArchive;
                if (archive is null) return null;
                var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
                return new RunArchiveContext(archive, "PY.RUN", key);
            }
            catch
            {
                return null;
            }
        }
    }

    /// <summary>Disposable wrapper that cancels (and disposes) a
    /// <see cref="CancellationTokenSource"/> exactly once.</summary>
    private sealed class CancelOnDispose : IDisposable
    {
        private CancellationTokenSource? _cts;

        public CancelOnDispose(CancellationTokenSource cts) { _cts = cts; }

        public void Dispose()
        {
            var cts = Interlocked.Exchange(ref _cts, null);
            if (cts is null) return;
            try { cts.Cancel(); } catch { /* best-effort */ }
            cts.Dispose();
        }
    }

    // -------------------------------------------------------------------------
    // Excel ↔ .NET boundary helpers
    // -------------------------------------------------------------------------

    /// <summary>
    /// Convert Excel-DNA's sentinel argument types into something
    /// <see cref="PyRun"/> understands.
    /// </summary>
    /// <remarks>
    /// Excel passes:
    /// <list type="bullet">
    ///   <item><see cref="ExcelMissing"/> when the argument was omitted —
    ///     map to <c>null</c> so PyRun calls the function with no positional
    ///     argument.</item>
    ///   <item><see cref="ExcelEmpty"/> when the cell is blank — also
    ///     <c>null</c>; the user's script sees no input.</item>
    ///   <item><see cref="ExcelError"/> when the referenced cell holds
    ///     an error value — propagate as <c>null</c> for now; surfacing
    ///     these as typed Python errors is a follow-up.</item>
    ///   <item><c>object[,]</c>, <c>string</c>, <c>double</c>, <c>bool</c>
    ///     pass through unchanged.</item>
    /// </list>
    /// </remarks>
    private static object? FromExcelArgument(object input)
    {
        if (input is ExcelMissing) return null;
        if (input is ExcelEmpty) return null;
        if (input is ExcelError) return null;
        return input;
    }

    private static object ToExcelOutput(object result)
    {
        return ReferenceEquals(result, PyRun.EmptyResult)
            ? ExcelEmpty.Value
            : result;
    }

    private static string ResolveFunctionName(object function)
    {
        if (function is string s && !string.IsNullOrWhiteSpace(s)) return s;
        return "transform";
    }
}
#endif
