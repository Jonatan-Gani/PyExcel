#if NETFRAMEWORK
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Drives the ribbon's Paste button: finds the newest archived run that
/// produced output, decodes <c>output.arrow</c> via
/// <see cref="ArrowMarshal"/>, and writes the result into the
/// user-configured target range.
///
/// <para><b>Threading / SAFE-1.</b> Same shape as
/// <see cref="ImportService"/>: the planner + range-resolve happen on the
/// main thread; the Arrow read + decode happen off it; the COM read of
/// the target range, the overwrite-confirmation prompt, and the write-
/// back are all queued via
/// <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/>.</para>
///
/// <para><b>Overwrite confirmation.</b> Before writing, the service
/// reads the target range's current <c>Value2</c> and asks
/// <see cref="PastePreflight.RangeHasContent"/> whether the paste would
/// destroy existing data. If so, a <see cref="MessageBox"/> with default
/// "No" is shown; the user must explicitly click "Yes" to proceed.
/// Cancelling logs to <see cref="LogDisplay"/> and aborts the paste
/// without touching the sheet.</para>
/// </summary>
public static class PasteService
{
    /// <summary>
    /// Execute the paste currently configured in <paramref name="state"/>.
    /// Returns immediately; the file read and the write-back happen
    /// asynchronously.
    /// </summary>
    public static void RunActivePaste(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        // --- main thread: plan -------------------------------------------
        PastePlan plan;
        try
        {
            var archive = PyExcelServices.RunArchive;
            if (archive is null)
            {
                Warn("Paste: the run archive service isn't initialised — " +
                     "no previous outputs to paste.");
                return;
            }
            var workbookKey = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            plan = PastePlanner.Create(state.PasteOutput, workbookKey, archive.List());
        }
        catch (FormatException fex)
        {
            Warn(fex.Message);
            return;
        }
        catch (Exception ex)
        {
            Fail($"Paste: failed to plan — {ex.Message}", ex);
            return;
        }

        // --- background thread: read + decode the archived output --------
        Task.Run(() =>
        {
            object? decoded;
            try
            {
                var bytes = File.ReadAllBytes(plan.SourceArrowPath);
                decoded = ArrowMarshal.Decode(bytes);
            }
            catch (FileNotFoundException)
            {
                Warn($"Paste: archived output '{plan.SourceArrowPath}' was " +
                     "deleted between planning and read.");
                return;
            }
            catch (Exception ex)
            {
                Fail($"Paste: failed to decode '{plan.SourceArrowPath}' — {ex.Message}", ex);
                return;
            }

            if (decoded is null)
            {
                Warn($"Paste: run {plan.SourceRunId} produced no payload to paste.");
                return;
            }

            if (decoded is ChartSpec or ChartImage)
            {
                // A chart isn't cell data — pasting it would write a
                // ToString() artefact. Charts render at run time via the
                // Run Python button; pointing the user there is the
                // actionable path.
                Warn($"Paste: run {plan.SourceRunId} produced a chart, which " +
                     "can't be pasted into cells — re-run the script with " +
                     "the Run Python button to render it.");
                return;
            }

            // --- main thread: preflight + write back ---------------------
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    var (rows, cols) = PastePreflight.Footprint(decoded);
                    if (rows == 0 || cols == 0)
                    {
                        Warn($"Paste: run {plan.SourceRunId} payload is empty.");
                        return;
                    }

                    dynamic app = ExcelDnaUtil.Application;
                    dynamic anchor = app.Range[plan.TargetRangeAddress];
                    dynamic targetRange = anchor.Resize[rows, cols];

                    // Read what's at the target right now. Excel-DNA's
                    // ExcelEmpty / ExcelMissing sentinels can appear in
                    // the COM hand-back for empty cells; strip them to
                    // null so the cross-platform PastePreflight (which
                    // doesn't reference those types) sees a clean
                    // snapshot.
                    object? before = StripExcelSentinels(targetRange.Value2);

                    if (PastePreflight.RangeHasContent(before)
                        && !ConfirmOverwrite(plan.TargetRangeAddress, rows, cols, plan.SourceRunId))
                    {
                        Trace.WriteLine(
                            $"Paste: user cancelled overwrite at " +
                            $"'{plan.TargetRangeAddress}'.");
                        LogDisplay.WriteLine(
                            $"Paste: cancelled — '{plan.TargetRangeAddress}' kept.");
                        return;
                    }

                    WriteToRange(targetRange, decoded);
                    Trace.WriteLine(
                        $"Paste: pasted run {plan.SourceRunId} into '{plan.TargetRangeAddress}'.");
                }
                catch (Exception ex)
                {
                    Fail(
                        $"Paste: failed to write into '{plan.TargetRangeAddress}' — {ex.Message}",
                        ex);
                }
            });
        });
    }

    /// <summary>Show the destructive-paste confirmation. Defaults to
    /// <see cref="MessageBoxDefaultButton.Button2"/> ("No") so an
    /// accidental Enter doesn't overwrite the user's data. Returns
    /// <see langword="true"/> iff the user clicked Yes.</summary>
    private static bool ConfirmOverwrite(string targetAddress, int rows, int cols, string runId)
    {
        var prompt =
            $"The target range '{targetAddress}' ({rows}×{cols}) contains " +
            $"values that will be overwritten by the paste from run {runId}." +
            Environment.NewLine + Environment.NewLine +
            "Continue?";

        var answer = MessageBox.Show(
            prompt,
            "PyExcel — confirm overwrite",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        return answer == DialogResult.Yes;
    }

    /// <summary>Recursively replace Excel-DNA's <see cref="ExcelEmpty"/>
    /// and <see cref="ExcelMissing"/> sentinels with <see langword="null"/>
    /// so the cross-platform preflight sees a clean snapshot.
    /// <see cref="ExcelError"/> is preserved — an error cell is content
    /// the user might still care about, and a paste over it is destructive
    /// in the same way as a paste over a number.</summary>
    private static object? StripExcelSentinels(object? value2)
    {
        if (value2 is null) return null;
        if (value2 is ExcelEmpty || value2 is ExcelMissing) return null;

        if (value2 is object[,] arr)
        {
            int r0 = arr.GetLowerBound(0);
            int c0 = arr.GetLowerBound(1);
            int height = arr.GetLength(0);
            int width = arr.GetLength(1);
            var cleaned = new object?[height, width];
            for (int i = 0; i < height; i++)
                for (int j = 0; j < width; j++)
                {
                    var cell = arr[r0 + i, c0 + j];
                    cleaned[i, j] = (cell is ExcelEmpty || cell is ExcelMissing)
                        ? null
                        : cell;
                }
            return cleaned;
        }

        return value2;
    }

    /// <summary>Write a decoded Arrow payload into the resolved target
    /// range. Mirrors the previous write semantics: a 2-D
    /// <c>object?[,]</c> sets <paramref name="targetRange"/>'s
    /// <c>Value2</c> directly (the caller has already sized the range to
    /// the payload's footprint), a 1-D <c>object?[]</c> writes as a row,
    /// a scalar drops into the top-left cell.</summary>
    private static void WriteToRange(dynamic targetRange, object decoded)
    {
        switch (decoded)
        {
            case object?[,] table:
                targetRange.Value2 = table;
                return;
            case object?[] vector:
            {
                var grid = new object?[1, vector.Length];
                for (int i = 0; i < vector.Length; i++) grid[0, i] = vector[i];
                targetRange.Value2 = grid;
                return;
            }
            default:
                targetRange.Value2 = decoded;
                return;
        }
    }

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
