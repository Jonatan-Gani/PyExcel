#if NETFRAMEWORK
using System;
using System.Diagnostics;
using ExcelDna.Integration;
using PyExcel.Kernel.Client;

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
///   <item>We hand the work to <see cref="ExcelAsyncUtil.Run"/>. The first
///     call kicks off a background task and returns <c>#N/A</c> immediately
///     so Excel's UI doesn't freeze (SAFE-1). When the task completes, the
///     cell auto-refreshes with the real result. Subsequent calls with the
///     same parameter set reuse the cached result.</item>
///   <item>On the worker thread, the sync core
///     (<see cref="RunSynchronously"/>) translates Excel-DNA's sentinel
///     argument types into plain .NET, calls
///     <see cref="PyRun.Execute"/> against
///     <see cref="KernelHost.Default"/>, and converts the result back to
///     something Excel can spill.</item>
/// </list>
///
/// <para>Failures surface as <see cref="ExcelError.ExcelErrorValue"/>
/// (<c>#VALUE!</c> in the cell) with the full Python traceback logged
/// to <see cref="Trace"/> so the user can see it in DebugView or the
/// Excel-DNA log window. A richer error UI — a per-script error pane or
/// a hover tooltip — is a Phase 4 follow-up.</para>
///
/// <para><em>Cancel:</em> Excel-DNA cancels the background task if the
/// formula changes or the workbook closes while the run is in flight,
/// but the kernel itself doesn't act on CANCEL frames yet — that's a
/// worker.py follow-up. Until then, an in-flight run completes even
/// after the host cancels.</para>
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
        // ExcelAsyncUtil.Run dispatches the work to a background thread,
        // returns #N/A immediately, and refreshes the cell when the
        // worker completes. The `parameters` array is used as the cache
        // key — identical inputs short-circuit to the cached result
        // rather than re-spawning a job.
        return ExcelAsyncUtil.Run(
            functionName: "PY.RUN",
            parameters: new object?[] { script, input, function },
            function: () => RunSynchronously(script, input, function));
    }

    /// <summary>
    /// The actual blocking work — runs on Excel-DNA's worker thread, not
    /// on the calc thread. Same error-translation contract as the sync
    /// UDF had before SAFE-1: <see cref="KernelException"/> and unhandled
    /// exceptions both surface as <c>#VALUE!</c>, with the diagnostic
    /// detail logged via <see cref="Trace"/>.
    /// </summary>
    private static object RunSynchronously(string script, object input, object function)
    {
        try
        {
            var functionName = ResolveFunctionName(function);
            var inp = FromExcelArgument(input);

            var result = PyRun.Execute(
                script: script,
                input: inp,
                kwargs: null,
                client: KernelHost.Default.Client,
                function: functionName);

            return ToExcelOutput(result);
        }
        catch (KernelException kex)
        {
            Trace.WriteLine(
                $"[PY.RUN] kernel error [{kex.Code}] {kex.PythonType}: {kex.Message}\n{kex.PythonTraceback}");
            return ExcelError.ExcelErrorValue;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"[PY.RUN] host error: {ex}");
            return ExcelError.ExcelErrorValue;
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
