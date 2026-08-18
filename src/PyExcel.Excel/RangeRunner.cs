#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.Bridge;
using PyExcel.Common.Logging;
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
    /// <param name="outputDisplay">Optional sink that shows the script's captured
    /// <c>print()</c> output to the user after a <em>successful</em> run, when the
    /// selected action's <see cref="RibbonAction.KeepOutputOpen"/> is true (the
    /// default) and the run actually printed something. Invoked on Excel's main
    /// thread with the captured text. A failed run is surfaced through
    /// <paramref name="errorDisplay"/> instead and is unaffected by this. When null,
    /// no output window is shown (output still goes to <see cref="LogDisplay"/>).</param>
    public static void RunActiveScript(
        WorkbookState state,
        Func<IRunProgressSink>? progressFactory = null,
        Action<string>? errorDisplay = null,
        Func<ListOrientation?>? orientationChooser = null,
        Action<string>? outputDisplay = null)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        // Whether to keep the script's console output on screen after a
        // successful run. Carried by the selected action; defaults to true for
        // an ad-hoc run with no action selected.
        bool keepOutputOpen = state.SelectedAction?.KeepOutputOpen ?? true;

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

        // Short correlation id stamped on every line of this run. Runs are
        // serialised through one kernel exchange, but a log read after the
        // fact interleaves them with workbook events and Setup output, so the
        // id is what lets a reader pick one run out of the file.
        var runId = Guid.NewGuid().ToString("N").Substring(0, 6);
        var runClock = Stopwatch.StartNew();
        Step(runId, $"start — script '{script}', action '{state.SelectedAction?.Name ?? "(none)"}'");
        Step(runId, $"input field: {(string.IsNullOrWhiteSpace(state.PyInput) ? "(empty)" : state.PyInput)}");
        Step(runId, $"output field: {(string.IsNullOrWhiteSpace(state.PyOutput) ? "(empty)" : state.PyOutput)}");

        // The Output field uses the same binding grammar as Input, so a
        // dict return can be routed key-by-key to its own range. Parsing it
        // also means a malformed Output is reported here rather than as a
        // raw COM error deep in the write-back.
        IReadOnlyList<RangeBinding> outputBindings;
        try
        {
            outputBindings = RibbonRangeParser.Parse(state.PyOutput);
        }
        catch (FormatException fex)
        {
            Warn($"Run Python: the Output field is malformed — {fex.Message}");
            return;
        }

        string? outputAddress = outputBindings.Count > 0 ? outputBindings[0].RangeText : null;

        // --- main thread: read inputs + capture workbook dir --------------
        List<object?> inputs;
        string? workbookDir;
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            workbookDir = ResolveWorkbookDir(app);

            inputs = new List<object?>(inputBindings.Count);
            for (var i = 0; i < inputBindings.Count; i++)
            {
                var binding = inputBindings[i];
                var value = ReadRange(app, binding.RangeText);
                inputs.Add(value);
                Step(runId,
                    $"input[{i}] '{binding.Name ?? "(auto-named)"}' "
                    + $"type={PyExcelTypes.WireName(binding.DeclaredType)} "
                    + $"range={binding.RangeText} -> {DescribeCells(value)}");
            }
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
            // Declared out here so the finally can unsubscribe regardless of
            // where the try exits.
            EventHandler<KernelOutputEventArgs>? captureHandler = null;
            var outputLock = new object();
            System.Text.StringBuilder? outputBuffer =
                keepOutputOpen && outputDisplay is not null
                    ? new System.Text.StringBuilder()
                    : null;
            try
            {
                // Capture the run's console output (stdout/stderr) for the
                // post-run output window. The supervisor drains the child's
                // streams on background threads, so accumulate under a lock.
                // Best-effort: a line still in flight when the run returns may
                // miss the snapshot — it's already in the log window too.
                if (outputBuffer is not null)
                {
                    // Capture a non-null local for the closure — the compiler
                    // can't prove the captured field stays non-null inside a
                    // lambda, and warnings are errors here.
                    System.Text.StringBuilder buffer = outputBuffer;
                    captureHandler = (_, ev) =>
                    {
                        lock (outputLock) buffer.AppendLine(ev.Text);
                    };
                    KernelHost.Default.Supervisor.OutputReceived += captureHandler;
                }

                for (var i = 0; i < outputBindings.Count; i++)
                {
                    var b = outputBindings[i];
                    Step(runId,
                        $"output[{i}] '{b.Name ?? "(positional)"}' "
                        + $"type={PyExcelTypes.WireName(b.DeclaredType)} range={b.RangeText}");
                }
                Step(runId,
                    $"dispatching to kernel — {inputs.Count} input(s), "
                    + $"{(progress is not null ? "async/cancellable" : "synchronous")}");
                var kernelClock = Stopwatch.StartNew();

                var result = progress is not null
                    ? PyRun.ExecuteManyAsync(
                        script: script,
                        inputs: inputs,
                        kwargs: null,
                        client: KernelHost.Default.Client,
                        workbookDirectory: workbookDir,
                        cancellationToken: runToken,
                        archive: archiveContext,
                        inputBindings: ToRunBindings(inputBindings),
                        outputBindings: ToRunBindings(outputBindings)).GetAwaiter().GetResult()
                    : PyRun.ExecuteMany(
                        script: script,
                        inputs: inputs,
                        kwargs: null,
                        client: KernelHost.Default.Client,
                        workbookDirectory: workbookDir,
                        archive: archiveContext,
                        inputBindings: ToRunBindings(inputBindings),
                        outputBindings: ToRunBindings(outputBindings));

                kernelClock.Stop();
                Step(runId,
                    $"kernel returned in {kernelClock.ElapsedMilliseconds} ms — {DescribeResult(result)}");

                // Snapshot the captured output before the finally unsubscribes.
                string capturedOutput = "";
                if (outputBuffer is not null)
                    lock (outputLock) capturedOutput = outputBuffer.ToString();

                // --- main thread: write the result back ------------------
                if (outputAddress is not null)
                {
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            if (result is PyRunOutputs named)
                            {
                                WriteNamedResults(
                                    named, outputBindings, outputAddress, orientationChooser, runId);
                            }
                            else
                            {
                                Step(runId, $"writing {DescribeResult(result)} -> {outputAddress}");
                                WriteResult(outputAddress, result, orientationChooser);
                            }
                            Step(runId, $"done in {runClock.ElapsedMilliseconds} ms");
                        }
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

                // Keep the script's console output on screen after a successful
                // run when the action opted in and it printed something. A failed
                // run is surfaced by the catch blocks below, so this path is
                // success-only by construction.
                if (capturedOutput.Length > 0)
                    ShowOutput(outputDisplay, capturedOutput);
            }
            catch (OperationCanceledException)
            {
                // The user cancelled from the progress dialog — expected, not a
                // failure. The progress form closes in the finally; don't pop an
                // error dialog, just note it in the log.
                Step(runId, $"cancelled by the user after {runClock.ElapsedMilliseconds} ms");
                Warn("Run Python: run cancelled.");
            }
            catch (KernelException kex)
            {
                // Test emptiness, not null. KernelException.PythonType is a
                // non-nullable string (its constructor coerces null to ""), and
                // null-testing it would set the compiler's null-state for the
                // expression to "maybe null" for every use after this one —
                // which is what made the KernelErrorRecord below fail CS8604.
                var pythonType = kex.PythonType.Length == 0 ? "(none)" : kex.PythonType;
                Step(runId,
                    $"kernel error after {runClock.ElapsedMilliseconds} ms — "
                    + $"code={kex.Code} type={pythonType}");
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
                if (captureHandler is not null)
                    KernelHost.Default.Supervisor.OutputReceived -= captureHandler;
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

    /// <summary>Show a successful run's captured console output in the injected
    /// sink, marshalled onto Excel's main thread (the output viewer is
    /// COM/WinForms-affine). No-op when no sink was supplied. Best-effort: the
    /// same output is always in <see cref="LogDisplay"/> regardless.</summary>
    private static void ShowOutput(Action<string>? outputDisplay, string text)
    {
        if (outputDisplay is null) return;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { outputDisplay(text); }
                catch { /* best-effort surface */ }
            });
        }
        catch { /* queueing itself failed — the output is still in LogDisplay */ }
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

    /// <summary>
    /// Project parsed ribbon bindings into the kernel's wire bindings.
    /// Returns <see langword="null"/> for an empty list so the request omits
    /// the meta key entirely — its absence is what keeps the legacy
    /// positional dispatch available to callers that never declared
    /// anything.
    /// </summary>
    private static IReadOnlyList<RunBinding>? ToRunBindings(
        IReadOnlyList<RangeBinding> bindings)
    {
        if (bindings.Count == 0) return null;

        var wire = new List<RunBinding>(bindings.Count);
        foreach (var b in bindings)
        {
            wire.Add(new RunBinding(
                b.Name, PyExcelTypes.WireName(b.DeclaredType), b.RangeText));
        }
        return wire;
    }

    /// <summary>
    /// Write each named result from a dict return to its own range.
    ///
    /// <para>A result whose key matches an Output binding goes to that
    /// binding's range. Anything left over — the script returned a key the
    /// Output field doesn't mention — is reported rather than dropped,
    /// naming both the unrouted keys and the ranges that were available,
    /// because silently discarding a computed result is precisely the
    /// failure this contract exists to remove.</para>
    /// </summary>
    private static void WriteNamedResults(
        PyRunOutputs named,
        IReadOnlyList<RangeBinding> outputBindings,
        string fallbackAddress,
        Func<ListOrientation?>? orientationChooser,
        string runId)
    {
        var unrouted = new List<string>();

        for (var i = 0; i < named.Outputs.Count; i++)
        {
            var output = named.Outputs[i];
            var address = ResolveOutputAddress(output.Name, i, outputBindings);

            if (address is null)
            {
                unrouted.Add(output.Name ?? $"#{i + 1}");
                Step(runId, $"result '{output.Name ?? $"#{i + 1}"}' has no output binding — not written");
                continue;
            }
            Step(runId,
                $"writing '{output.Name ?? $"#{i + 1}"}' ({DescribeResult(output.Value)}) -> {address}");
            WriteResult(address, output.Value, orientationChooser);
        }

        if (unrouted.Count == 0) return;

        var available = outputBindings.Count == 0
            ? fallbackAddress
            : string.Join("; ", outputBindings.Select(b => b.Name is null
                ? b.RangeText
                : $"{b.Name}={b.RangeText}"));

        Warn(
            "Run Python: the script returned "
            + string.Join(", ", unrouted.Select(n => $"'{n}'"))
            + " but the Output field has nowhere to put "
            + (unrouted.Count == 1 ? "it" : "them")
            + $". Output currently reads: {available}. Add a binding named for each "
            + "returned key, e.g. 'total=Sheet1!A1; table=Sheet1!C1'.");
    }

    /// <summary>
    /// Pick the range for one named result: an Output binding of the same
    /// name wins; failing that, an unnamed binding in the same ordinal
    /// position is used, which is what makes a single-output script work
    /// without the user naming anything.
    /// </summary>
    private static string? ResolveOutputAddress(
        string? name, int index, IReadOnlyList<RangeBinding> outputBindings)
    {
        if (name is not null)
        {
            foreach (var b in outputBindings)
                if (string.Equals(b.Name, name, StringComparison.Ordinal))
                    return b.RangeText;
        }

        if (index < outputBindings.Count && outputBindings[index].Name is null)
            return outputBindings[index].RangeText;

        return null;
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

    /// <summary>
    /// The persistent log. Until this existed the run path wrote only to
    /// <see cref="Trace"/> and Excel-DNA's <c>LogDisplay</c> — neither of which
    /// reaches <c>%TEMP%\PyExcel_Debug.log</c> — so a completed run left no
    /// record on disk at all and a failed one left only whatever the user
    /// happened to still have on screen.
    /// </summary>
    private static readonly ILog Log = new FileLog();

    /// <summary>
    /// One step of a run, correlated by <paramref name="runId"/>.
    ///
    /// <para>File-log only, deliberately: this is the play-by-play a
    /// post-mortem needs, and pushing it to <c>LogDisplay</c> too would bury
    /// the warnings the user is meant to read in a wall of detail.</para>
    /// </summary>
    private static void Step(string runId, string message)
        => Log.Info($"run {runId}: {message}");

    private static void Warn(string message)
    {
        Trace.WriteLine(message);
        Log.Warn(message);
        LogDisplay.WriteLine(message);
    }

    private static void Fail(string message, Exception ex)
    {
        Trace.WriteLine(message);
        Trace.WriteLine(ex.ToString());
        Log.Error(message, ex);
        LogDisplay.WriteLine(message);
    }

    /// <summary>Row x column description of a value read off a range, for the
    /// trace. Mirrors how the kernel will see it once decoded to a grid.</summary>
    private static string DescribeCells(object? value) => value switch
    {
        null => "empty",
        object[,] grid => $"{grid.GetLength(0)}x{grid.GetLength(1)} cells",
        _ => "1x1 cell",
    };

    /// <summary>Shape description of a decoded result, for the trace.</summary>
    private static string DescribeResult(object result)
    {
        if (ReferenceEquals(result, PyRun.EmptyResult)) return "None";
        return result switch
        {
            PyRunOutputs named => $"{named.Outputs.Count} named result(s)",
            object[,] table => $"table {table.GetLength(0)}x{table.GetLength(1)}",
            ChartSpec => "chart spec",
            ChartImage => "chart image",
            _ => $"scalar ({result.GetType().Name})",
        };
    }
}
#endif
