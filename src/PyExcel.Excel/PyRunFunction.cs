#if NETFRAMEWORK
using System;
using System.Diagnostics;
using ExcelDna.Integration;
using PyExcel.Kernel.Client;

namespace PyExcel.Excel;

/// <summary>
/// The <c>=PY.RUN</c> worksheet function — the smallest possible Excel-DNA
/// surface on top of <see cref="PyRun"/>.
///
/// <para>The flow:</para>
///
/// <list type="number">
///   <item>Excel passes a script path and one input argument
///     (cell / range / array / scalar) plus optional function name.</item>
///   <item>We translate Excel-DNA's sentinel argument types
///     (<see cref="ExcelMissing"/>, <see cref="ExcelEmpty"/>,
///     <see cref="ExcelError"/>) into <c>null</c> so
///     <see cref="PyRun"/> sees a regular .NET shape.</item>
///   <item><see cref="PyRun.Execute"/> handles the marshal + dispatch
///     against the process-wide <see cref="KernelHost.Default"/>.</item>
///   <item>The result is translated back to something Excel can spill:
///     <see cref="PyRun.EmptyResult"/> becomes <see cref="ExcelEmpty.Value"/>;
///     anything else passes through unchanged.</item>
/// </list>
///
/// <para>Failures surface as <see cref="ExcelError.ExcelErrorValue"/>
/// (<c>#VALUE!</c> in the cell) with the full Python traceback logged
/// to <see cref="Debug"/> / <see cref="Trace"/> so the user can see it
/// in DebugView or the Excel-DNA log window. A richer error UI — a
/// per-script error pane or a hover tooltip — is a Phase 4 follow-up.</para>
///
/// <para><em>SAFE-1 note:</em> this UDF is currently synchronous; it
/// blocks the Excel calc thread for the duration of the run. The
/// roadmap calls for an async variant that returns
/// <see cref="ExcelAsyncUtil"/>-driven results so long jobs don't
/// freeze the UI. That's a separate Phase 4 item.</para>
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
            // The kernel-side traceback is the actually-useful diagnostic.
            // We log it for DebugView / the Excel-DNA log window; the cell
            // gets #VALUE! since a string return would look like a result.
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
