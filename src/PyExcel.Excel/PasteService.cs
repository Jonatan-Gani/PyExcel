#if NETFRAMEWORK
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
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
/// main thread; the Arrow read + decode happen off it; the COM write-back
/// is queued via
/// <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/>.</para>
///
/// <para><b>Overwrite confirmation</b> is intentionally not in this
/// service. The roadmap calls for it on the Paste flow specifically —
/// because pasting into a populated range is destructive — but the
/// confirmation dialog is WinForms (Phase 8). For now the paste
/// overwrites; the run-archive retains the destination data only
/// indirectly (via the next run that would land on the same cells), so
/// the Phase-8 dialog is the right place to add the prompt without
/// duplicating data.</para>
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

            // --- main thread: write back into the target range -----------
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    WriteToRange(plan.TargetRangeAddress, decoded);
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

    /// <summary>Write a decoded Arrow payload into the target range.
    /// Mirrors <see cref="RangeRunner"/>'s write semantics: a 2-D
    /// <c>object?[,]</c> resizes the anchor to its dimensions, a 1-D
    /// <c>object?[]</c> writes as a row (we don't have orientation
    /// metadata for the archived buffer beyond what ArrowMarshal recovers
    /// — table is the default), a scalar drops into the top-left
    /// cell.</summary>
    private static void WriteToRange(string targetAddress, object decoded)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic anchor = app.Range[targetAddress];

        switch (decoded)
        {
            case object?[,] table:
            {
                int rows = table.GetLength(0);
                int cols = table.GetLength(1);
                if (rows == 0 || cols == 0) return;
                dynamic target = anchor.Resize[rows, cols];
                target.Value2 = table;
                return;
            }
            case object?[] vector:
            {
                if (vector.Length == 0) return;
                // ArrowMarshal preserves vector orientation via metadata;
                // its decode returns a 1-D array which we spill as a row.
                // A column-vector caller (rare from the archive path
                // where outputs are usually tables or scalars) can still
                // transpose at the cell level.
                dynamic target = anchor.Resize[1, vector.Length];
                var grid = new object?[1, vector.Length];
                for (int i = 0; i < vector.Length; i++) grid[0, i] = vector[i];
                target.Value2 = grid;
                return;
            }
            default:
                anchor.Value2 = decoded;
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
