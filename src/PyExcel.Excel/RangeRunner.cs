#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.Kernel.Client;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Drives the ribbon's <c>Run Python</c> button: reads the configured
/// input ranges off the active sheet, runs the selected script through
/// the kernel, and writes the result back to the output range.
///
/// <para>This is the COM-bound counterpart to the <c>=PY.RUN</c> UDF.
/// Where the UDF takes a single argument and spills its result into the
/// calling cell, the button reads the Python-group fields the user
/// configured in the ribbon (<see cref="WorkbookState.SelectedScript"/>,
/// <see cref="WorkbookState.PyInput"/>,
/// <see cref="WorkbookState.PyOutput"/>) and writes into an explicit
/// output range.</para>
///
/// <para><b>Threading / SAFE-1.</b> Range reads and writes are COM calls
/// that must run on Excel's main thread; the kernel exchange must
/// <em>not</em> block it. So the flow is:</para>
/// <list type="number">
///   <item>On the caller's thread (the ribbon callback, which Excel
///     invokes on the main thread) — resolve and read every input range
///     into a plain managed array, and capture the script / output /
///     workbook-dir as strings. No pipe traffic here.</item>
///   <item>On a background <see cref="Task"/> — encode, run the kernel
///     exchange, decode. This is the only part that can block.</item>
///   <item>Back on the main thread via
///     <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/> — write
///     the decoded result into the output range.</item>
/// </list>
///
/// <para>All COM access is late-bound through <c>dynamic</c> on
/// <see cref="ExcelDnaUtil.Application"/>, so this assembly needs no
/// Office PIA reference (which would also break the PIA-less Windows CI
/// build). Errors surface to <see cref="Trace"/> and Excel-DNA's
/// <see cref="LogDisplay"/>, matching <see cref="PyRunFunction"/>.</para>
/// </summary>
public static class RangeRunner
{
    /// <summary>
    /// Execute the script currently selected in <paramref name="state"/>
    /// against its configured input/output ranges. Returns immediately;
    /// the kernel run and the write-back happen asynchronously.
    /// </summary>
    /// <remarks>
    /// Must be called on Excel's main thread (a ribbon callback satisfies
    /// this) because it reads ranges synchronously before handing off.
    /// </remarks>
    public static void RunActiveScript(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        if (string.IsNullOrWhiteSpace(state.SelectedScript))
        {
            Warn("Run Python: no script is selected. Pick a script in the ribbon first.");
            return;
        }

        IReadOnlyList<RangeBinding> inputBindings;
        try
        {
            inputBindings = RibbonRangeParser.Parse(state.PyInput);
        }
        catch (FormatException fex)
        {
            Warn($"Run Python: the Input field is malformed — {fex.Message}");
            return;
        }

        string script = state.SelectedScript!;
        string? outputAddress = string.IsNullOrWhiteSpace(state.PyOutput) ? null : state.PyOutput;

        // --- main thread: read inputs + capture workbook dir --------------
        List<object?> inputs;
        string? workbookDir;
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            workbookDir = ResolveWorkbookDir(app);

            inputs = new List<object?>(inputBindings.Count);
            foreach (var binding in inputBindings)
                inputs.Add(ReadRange(app, binding.RangeText));
        }
        catch (Exception ex)
        {
            Fail($"Run Python: failed to read the input range(s) — {ex.Message}", ex);
            return;
        }

        var archiveContext = BuildArchiveContext();

        // --- background thread: the kernel exchange (may block) -----------
        Task.Run(() =>
        {
            try
            {
                var result = PyRun.ExecuteMany(
                    script: script,
                    inputs: inputs,
                    kwargs: null,
                    client: KernelHost.Default.Client,
                    workbookDirectory: workbookDir,
                    archive: archiveContext);

                // --- main thread: write the result back ------------------
                if (outputAddress is null)
                {
                    // Nothing to write into; the script presumably has side
                    // effects or the user only wanted to exercise it.
                    return;
                }
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { WriteResult(outputAddress, result); }
                    catch (Exception ex)
                    {
                        Fail($"Run Python: failed to write the output range — {ex.Message}", ex);
                    }
                });
            }
            catch (KernelException kex)
            {
                var record = new KernelErrorRecord(
                    Timestamp: DateTimeOffset.UtcNow,
                    Source: "Run Python button",
                    Code: kex.Code,
                    PythonType: kex.PythonType,
                    Message: kex.Message,
                    PythonTraceback: kex.PythonTraceback,
                    ScriptPath: script);
                RecordError(record);
                Fail(record.FormatForClipboard(), kex);
            }
            catch (Exception ex)
            {
                var record = new KernelErrorRecord(
                    Timestamp: DateTimeOffset.UtcNow,
                    Source: "Run Python button",
                    Code: "HostError",
                    PythonType: ex.GetType().Name,
                    Message: ex.Message,
                    PythonTraceback: ex.ToString(),
                    ScriptPath: script);
                RecordError(record);
                Fail(record.FormatForClipboard(), ex);
            }
        });
    }

    /// <summary>Best-effort push into the per-workbook last-error slot.
    /// A failure here must not stop the user-facing error from reaching
    /// LogDisplay / Trace.</summary>
    private static void RecordError(KernelErrorRecord record)
    {
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            PyExcelServices.Errors.Record(key, record);
        }
        catch
        {
            // Best-effort.
        }
    }

    /// <summary>Build the archive context for this ribbon-button run.
    /// Best-effort: archiving is diagnostic, not load-bearing, so if the
    /// service slot or the workbook lookup throws we just return null and
    /// the run proceeds unarchived.</summary>
    private static RunArchiveContext? BuildArchiveContext()
    {
        try
        {
            var archive = PyExcelServices.RunArchive;
            if (archive is null) return null;
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            return new RunArchiveContext(archive, "Run Python button", key);
        }
        catch
        {
            return null;
        }
    }

    // -------------------------------------------------------------------------
    // COM helpers — late-bound, main-thread only
    // -------------------------------------------------------------------------

    /// <summary>Workbook directory of the active workbook, or
    /// <see langword="null"/> for an unsaved workbook (empty Path).
    /// Used to resolve relative script paths.</summary>
    private static string? ResolveWorkbookDir(dynamic app)
    {
        dynamic wb = app.ActiveWorkbook;
        if (wb is null) return null;
        string path = (string)wb.Path;
        return string.IsNullOrEmpty(path) ? null : path;
    }

    /// <summary>Read a range address into a plain managed value:
    /// <c>object?[,]</c> (0-based) for a multi-cell range, or a scalar
    /// for a single cell. Excel hands back a 1-based 2-D array for ranges;
    /// we re-base it to 0 so <see cref="ArrowMarshal"/> can index it
    /// directly.</summary>
    private static object? ReadRange(dynamic app, string address)
    {
        dynamic range = app.Range[address];
        object value = range.Value2;
        return NormalizeFromExcel(value);
    }

    private static object? NormalizeFromExcel(object? value)
    {
        if (value is object[,] arr)
        {
            int r0 = arr.GetLowerBound(0);
            int c0 = arr.GetLowerBound(1);
            int rows = arr.GetLength(0);
            int cols = arr.GetLength(1);
            var norm = new object?[rows, cols];
            for (var i = 0; i < rows; i++)
                for (var j = 0; j < cols; j++)
                    norm[i, j] = arr[r0 + i, c0 + j];
            return norm;
        }
        // Single cell: Value2 is already a scalar (double / string / bool /
        // null), which ArrowMarshal.EncodeScalar handles directly.
        return value;
    }

    /// <summary>Write a decoded result into the output range. A table
    /// result resizes the anchor to its dimensions; a scalar drops into
    /// the top-left cell. A <c>None</c> return writes nothing.</summary>
    private static void WriteResult(string outputAddress, object result)
    {
        if (ReferenceEquals(result, PyRun.EmptyResult)) return;

        dynamic app = ExcelDnaUtil.Application;
        dynamic anchor = app.Range[outputAddress];

        if (result is object?[,] table)
        {
            int rows = table.GetLength(0);
            int cols = table.GetLength(1);
            if (rows == 0 || cols == 0) return;
            // Resize the anchor's top-left to the result's footprint so the
            // whole block lands even if the user only named a single cell.
            dynamic target = anchor.Resize[rows, cols];
            target.Value2 = table;
        }
        else
        {
            anchor.Value2 = result;
        }
    }

    // -------------------------------------------------------------------------
    // Diagnostics
    // -------------------------------------------------------------------------

    private static void Warn(string message)
    {
        Trace.WriteLine(message);
        LogDisplay.WriteLine(message);
    }

    private static void Fail(string message, Exception ex)
    {
        Trace.WriteLine(message);
        Trace.WriteLine(ex.ToString());
        LogDisplay.WriteLine(message);
    }
}
#endif
