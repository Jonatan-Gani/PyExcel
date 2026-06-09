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
/// Drives the ribbon's Import button: reads the user-configured CSV /
/// TSV file off disk and writes its contents into the configured target
/// range. The cross-platform pieces — field validation, path resolution,
/// delimiter detection, line parsing — live in <see cref="ImportPlanner"/>
/// and <see cref="CsvParser"/>; this class is the COM-bound shell that
/// dispatches them on the right threads.
///
/// <para><b>Threading / SAFE-1.</b> File I/O happens off the main thread;
/// COM access happens on it (queued via
/// <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/>). The ribbon
/// callback returns immediately so Excel's UI thread stays responsive
/// even on a multi-megabyte CSV import.</para>
///
/// <para>Per-cell type inference matches Excel's built-in CSV import
/// (and what users expect from a CSV): a string that parses as a
/// <see cref="double"/> under
/// <see cref="System.Globalization.CultureInfo.InvariantCulture"/>
/// becomes a numeric cell; <c>"TRUE"</c> / <c>"FALSE"</c>
/// (case-insensitive) become boolean cells; everything else stays a
/// string. Leading-zero strings (<c>"00123"</c>) intentionally fall
/// through to string-typed cells because parsing destroys the leading
/// zeros — the user's typing is the signal, the inference is a
/// convenience.</para>
/// </summary>
public static class ImportService
{
    /// <summary>
    /// Execute the import currently configured in <paramref name="state"/>.
    /// Returns immediately; the file read and the write-back happen
    /// asynchronously.
    /// </summary>
    public static void RunActiveImport(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        // --- main thread: plan + capture workbook dir --------------------
        ImportPlan plan;
        try
        {
            string? workbookDir = ResolveWorkbookDir();
            plan = ImportPlanner.Create(state.ImportInput, state.ImportOutput, workbookDir);
        }
        catch (FormatException fex)
        {
            Warn(fex.Message);
            return;
        }
        catch (Exception ex)
        {
            Fail($"Import: failed to plan — {ex.Message}", ex);
            return;
        }

        // --- background thread: file read + parse ------------------------
        Task.Run(() =>
        {
            object?[,] table;
            try
            {
                using var stream = new FileStream(
                    plan.AbsoluteSourcePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read);
                var rows = CsvParser.Parse(stream, plan.Delimiter);
                table = ToObjectGrid(rows);
            }
            catch (FileNotFoundException)
            {
                Warn($"Import: file not found at '{plan.AbsoluteSourcePath}'.");
                return;
            }
            catch (DirectoryNotFoundException)
            {
                Warn($"Import: directory not found for '{plan.AbsoluteSourcePath}'.");
                return;
            }
            catch (Exception ex)
            {
                Fail($"Import: failed to read '{plan.AbsoluteSourcePath}' — {ex.Message}", ex);
                return;
            }

            if (table.GetLength(0) == 0 || table.GetLength(1) == 0)
            {
                Warn($"Import: '{plan.AbsoluteSourcePath}' parsed to zero rows.");
                return;
            }

            // --- main thread: write the parsed table back ----------------
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { WriteTable(plan.TargetRangeAddress, table); }
                catch (Exception ex)
                {
                    Fail(
                        $"Import: failed to write into '{plan.TargetRangeAddress}' — {ex.Message}",
                        ex);
                }
            });
        });
    }

    /// <summary>Convert the parser's ragged record list into a
    /// rectangular <c>object?[,]</c>. Short rows are padded with nulls so
    /// every column lands; rows can't be longer than the widest because
    /// the width is the max field count across the file.</summary>
    private static object?[,] ToObjectGrid(System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<string>> rows)
    {
        int height = rows.Count;
        int width = 0;
        for (int i = 0; i < height; i++)
            if (rows[i].Count > width) width = rows[i].Count;

        var grid = new object?[height, width];
        for (int r = 0; r < height; r++)
        {
            var row = rows[r];
            int cols = row.Count;
            for (int c = 0; c < width; c++)
            {
                grid[r, c] = c < cols ? CsvCellTypeInference.Infer(row[c]) : null;
            }
        }
        return grid;
    }

    /// <summary>Workbook directory of the active workbook, or
    /// <see langword="null"/> for an unsaved workbook (empty Path).
    /// Mirrors <see cref="RangeRunner"/>'s helper of the same name —
    /// kept private rather than shared so each service can be read end
    /// to end without jumping files.</summary>
    private static string? ResolveWorkbookDir()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb is null) return null;
            string path = (string)wb.Path;
            return string.IsNullOrEmpty(path) ? null : path;
        }
        catch
        {
            // Live host can fail to surface ActiveWorkbook in odd states
            // (no workbook, just the splash, …). Treat as "no dir."
            return null;
        }
    }

    /// <summary>Resize the target-range anchor to the data dimensions and
    /// write the typed grid. Excel's <c>Value2</c> accepts
    /// <c>object[,]</c> directly; the 1-based array indices Excel itself
    /// uses are internal to the COM call — passing a 0-based managed
    /// array works.</summary>
    private static void WriteTable(string targetAddress, object?[,] table)
    {
        int rows = table.GetLength(0);
        int cols = table.GetLength(1);
        if (rows == 0 || cols == 0) return;

        dynamic app = ExcelDnaUtil.Application;
        dynamic anchor = app.Range[targetAddress];
        dynamic target = anchor.Resize[rows, cols];
        target.Value2 = table;
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
