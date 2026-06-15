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
    /// <param name="state">The active workbook state.</param>
    /// <param name="progressFactory">Optional factory creating a modeless
    /// progress sink (the WinForms <c>ProgressForm</c>). When supplied the
    /// run goes through the async, cancellable path and forwards the
    /// kernel's <c>PROGRESS</c> frames to the sink; when null the original
    /// synchronous path runs unchanged.</param>
    /// <param name="errorDisplay">Optional sink that shows a run failure to
    /// the user in a modal, dismiss-to-continue dialog. Invoked on Excel's
    /// main thread with the formatted error block when a run fails (a Python
    /// traceback, a read/write fault, …) so the failure stays on screen until
    /// the user closes it instead of scrolling past in the log window. When
    /// null, failures are only written to <see cref="LogDisplay"/> as before.</param>
    /// <param name="orientationChooser">Optional callback (invoked on Excel's main
    /// thread) asked which way a 1-D list result should spill when the Output range
    /// is a single cell, where the direction is ambiguous. Returns null to cancel the
    /// write. When null (no UI), a 1-D list into a single cell defaults to a row.</param>
    public static void RunActiveScript(
        WorkbookState state,
        Func<IRunProgressSink>? progressFactory = null,
        Action<string>? errorDisplay = null,
        Func<ListOrientation?>? orientationChooser = null)
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

        // Resolve the selected script (a bare userScripts filename like "foo.py")
        // to its real path under the project's userScripts folder, so the kernel
        // loads the right file. The project root is the dedicated folder chosen on
        // Enable (state.ProjectDir) if set, else the workbook-derived default —
        // the same rule the ribbon's Edit-script and KernelHost use. An
        // already-rooted path (legacy / power user) is passed through untouched.
        if (!System.IO.Path.IsPathRooted(script))
        {
            var projectDir = !string.IsNullOrEmpty(state.ProjectDir)
                ? state.ProjectDir
                : PyExcel.Common.ProjectDirectory.Resolve(workbookDir);
            if (!string.IsNullOrEmpty(projectDir))
                script = System.IO.Path.Combine(projectDir!, "userScripts", script);
        }

        var archiveContext = BuildArchiveContext();

        // Optional progress UI: created on the main thread (shows a modeless
        // dialog) and fed the kernel's PROGRESS frames. When absent the run
        // takes the original synchronous, non-cancellable path.
        IRunProgressSink? progress = progressFactory?.Invoke();
        EventHandler<ProgressReceivedEventArgs>? progressHandler = null;
        if (progress is not null)
        {
            IRunProgressSink sink = progress;
            progressHandler = (_, ev) => sink.Report(ev.Percent, ev.Message);
            KernelHost.Default.Client.ProgressReceived += progressHandler;
        }
        var runToken = progress?.CancellationToken ?? default;

        // --- background thread: the kernel exchange (may block) -----------
        Task.Run(() =>
        {
            try
            {
                var result = progress is not null
                    ? PyRun.ExecuteManyAsync(
                        script: script,
                        inputs: inputs,
                        kwargs: null,
                        client: KernelHost.Default.Client,
                        workbookDirectory: workbookDir,
                        cancellationToken: runToken,
                        archive: archiveContext).GetAwaiter().GetResult()
                    : PyRun.ExecuteMany(
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
                    try { WriteResult(outputAddress, result, orientationChooser); }
                    catch (Exception ex)
                    {
                        var message = $"Run Python: failed to write the output range — {ex.Message}";
                        Fail(message, ex);
                        // Already on the main thread here (QueueAsMacro), so show
                        // the modal inline rather than re-queuing it.
                        try { errorDisplay?.Invoke(message); } catch { /* best-effort */ }
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
                var block = record.FormatForClipboard();
                Fail(block, kex);
                ShowError(errorDisplay, block);
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
                var block = record.FormatForClipboard();
                Fail(block, ex);
                ShowError(errorDisplay, block);
            }
            finally
            {
                if (progressHandler is not null)
                    KernelHost.Default.Client.ProgressReceived -= progressHandler;
                progress?.Complete();
            }
        });
    }

    /// <summary>Show a run failure in the injected modal error sink, marshalled
    /// onto Excel's main thread (the kernel catch blocks run on the background
    /// task, but the dialog is COM/WinForms-affine and must be owned by Excel's
    /// window). No-op when no sink was supplied. Best-effort: surfacing the
    /// error must never throw on top of the failure that triggered it.</summary>
    private static void ShowError(Action<string>? errorDisplay, string message)
    {
        if (errorDisplay is null) return;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { errorDisplay(message); }
                catch { /* best-effort surface */ }
            });
        }
        catch { /* queueing itself failed — the error is still in LogDisplay */ }
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

    /// <summary>Write a decoded result into the output range. A 2-D table resizes
    /// the anchor to its dimensions; a 1-D list spills as a row or column (see
    /// <see cref="WriteVector"/>); a scalar drops into the top-left cell. A
    /// <c>None</c> return writes nothing. A chart spec builds a native chart anchored
    /// at the output range; a chart image embeds as a picture there.</summary>
    private static void WriteResult(
        string outputAddress, object result, Func<ListOrientation?>? orientationChooser)
    {
        if (ReferenceEquals(result, PyRun.EmptyResult)) return;

        dynamic app = ExcelDnaUtil.Application;
        dynamic anchor = app.Range[outputAddress];

        if (result is ChartSpec chartSpec)
        {
            // Parse before touching the sheet so a malformed spec fails
            // with a clean message and no orphan ChartObject.
            var document = ChartSpecParser.Parse(chartSpec.Json);

            double left = (double)anchor.Left;
            double top = (double)anchor.Top;
            // A multi-cell output range doubles as the chart's footprint;
            // a single-cell anchor gets the v1 default size.
            double width = (double)anchor.Width;
            double height = (double)anchor.Height;
            if ((int)anchor.Cells.Count <= 1)
            {
                width = 800;
                height = 500;
            }
            ChartBuilder.Build(anchor.Worksheet, document, left, top, width, height);
            return;
        }

        if (result is ChartImage chartImage)
        {
            ChartBuilder.EmbedImage(
                anchor.Worksheet, chartImage, (double)anchor.Left, (double)anchor.Top);
            return;
        }

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
        else if (result is object?[] vector)
        {
            WriteVector(anchor, vector, orientationChooser);
        }
        else
        {
            anchor.Value2 = result;
        }
    }

    /// <summary>Spill a 1-D list result into the output range, choosing its
    /// direction from the range's shape (<see cref="OrientationResolver"/>):
    /// <list type="bullet">
    ///   <item>a single row or column fixes the direction;</item>
    ///   <item>a single cell is ambiguous, so the injected
    ///     <paramref name="orientationChooser"/> asks the user — defaulting to a row
    ///     when no chooser is wired, and aborting the write (leaving the sheet
    ///     untouched) when the user dismisses the prompt;</item>
    ///   <item>a 2-D block target is rejected — a list can't fill it unambiguously —
    ///     by throwing, which the caller surfaces as a clean run-failure message.</item>
    /// </list>
    /// The list spills its own length from the anchor's top-left in the chosen
    /// direction; the target range only disambiguates which way.</summary>
    private static void WriteVector(
        dynamic anchor, object?[] vector, Func<ListOrientation?>? orientationChooser)
    {
        if (vector.Length == 0) return;

        int rows = (int)anchor.Rows.Count;
        int cols = (int)anchor.Columns.Count;
        var resolution = OrientationResolver.Resolve(rows, cols);

        if (resolution.IsInvalid)
            throw new InvalidOperationException(
                "the result is a 1-D list but the Output range is a 2-D block. " +
                "Point Output at a single cell, a single row, or a single column.");

        ListOrientation orientation;
        if (resolution.Ask)
        {
            var chosen = orientationChooser?.Invoke();
            if (orientationChooser is not null && chosen is null)
            {
                // The user dismissed the direction prompt — leave the sheet alone.
                Warn("Run Python: write cancelled — no list direction chosen.");
                return;
            }
            orientation = chosen ?? ListOrientation.Horizontal;
        }
        else
        {
            orientation = resolution.Orientation;
        }

        object?[,] grid = orientation == ListOrientation.Vertical
            ? ToColumn(vector)
            : ToRow(vector);
        dynamic target = anchor.Resize[grid.GetLength(0), grid.GetLength(1)];
        target.Value2 = grid;
    }

    /// <summary>Reshape a 1-D vector into a 1×N row.</summary>
    private static object?[,] ToRow(object?[] vector)
    {
        var grid = new object?[1, vector.Length];
        for (int i = 0; i < vector.Length; i++) grid[0, i] = vector[i];
        return grid;
    }

    /// <summary>Reshape a 1-D vector into an N×1 column.</summary>
    private static object?[,] ToColumn(object?[] vector)
    {
        var grid = new object?[vector.Length, 1];
        for (int i = 0; i < vector.Length; i++) grid[i, 0] = vector[i];
        return grid;
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
