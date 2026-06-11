#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Drives the ribbon's Export button: reads the configured source range
/// off the active sheet and writes it as CSV / TSV to the user-typed
/// destination file. Sibling of <see cref="ImportService"/>; same
/// SAFE-1 threading contract (COM on the main thread, file I/O off it),
/// same diagnostic surface (LogDisplay + Trace).
///
/// <para>Cell-value formatting (string conversion before CSV quoting) is
/// delegated to <see cref="CsvCellFormatter"/>: numbers use the
/// invariant-culture round-trip format so re-importing is loss-free;
/// booleans become <c>"TRUE"</c> / <c>"FALSE"</c> (Excel's display
/// convention, also what <see cref="CsvCellTypeInference"/> recognises
/// on round-trip); <see cref="DateTime"/> uses ISO 8601; nulls and the
/// Excel-DNA <see cref="ExcelEmpty"/> / <see cref="ExcelError"/>
/// sentinels render as the empty string. The quoting itself
/// (delimiter / quote / newline → wrap) lives in <see cref="CsvWriter"/>.</para>
/// </summary>
public static class ExportService
{
    /// <summary>
    /// Execute the export currently configured in <paramref name="state"/>.
    /// Returns immediately; the range read happens on the calling
    /// (main) thread, file I/O happens on a background task.
    /// </summary>
    public static void RunActiveExport(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        // --- main thread: plan + read the range --------------------------
        ExportPlan plan;
        string?[,] rows;
        try
        {
            string? workbookDir = ResolveWorkbookDir();
            plan = ExportPlanner.Create(state.ExportInput, state.ExportOutput, workbookDir);
            rows = ReadAsStringGrid(plan.SourceRangeAddress);
        }
        catch (FormatException fex)
        {
            Warn(fex.Message);
            return;
        }
        catch (Exception ex)
        {
            Fail($"Export: failed to read the source range — {ex.Message}", ex);
            return;
        }

        int height = rows.GetLength(0);
        int width = rows.GetLength(1);
        if (height == 0 || width == 0)
        {
            Warn($"Export: '{plan.SourceRangeAddress}' resolved to an empty range.");
            return;
        }

        // --- background thread: file write -------------------------------
        Task.Run(() =>
        {
            try
            {
                EnsureDirectory(plan.AbsoluteTargetPath);
                using var stream = new FileStream(
                    plan.AbsoluteTargetPath,
                    FileMode.Create,
                    FileAccess.Write,
                    FileShare.Read);
                CsvWriter.Write(
                    stream,
                    EnumerateRows(rows),
                    delimiter: plan.Delimiter,
                    lineTerminator: "\r\n",
                    encoding: null,
                    writeBom: false);
                Trace.WriteLine($"Export: wrote {height}×{width} to '{plan.AbsoluteTargetPath}'.");
            }
            catch (Exception ex)
            {
                Fail($"Export: failed to write '{plan.AbsoluteTargetPath}' — {ex.Message}", ex);
            }
        });
    }

    /// <summary>Run a batch of exports (the Export Wizard). Reads every
    /// source range on the main thread, then writes each file on a
    /// background thread — the same per-export pipeline as
    /// <see cref="RunActiveExport"/>, looped. Each job's planning reuses
    /// <see cref="ExportPlanner"/>; the wizard has already validated them,
    /// so a plan failure here is surfaced and the row skipped.</summary>
    public static void RunBatch(IReadOnlyList<ExportJob> jobs, string? workbookDirectory)
    {
        if (jobs is null) throw new ArgumentNullException(nameof(jobs));
        if (jobs.Count == 0) { Warn("Export Wizard: no rows to export."); return; }

        // --- main thread: plan + read every source range ----------------
        var planned = new List<(ExportPlan Plan, string?[,] Rows)>(jobs.Count);
        try
        {
            foreach (var job in jobs)
            {
                var plan = ExportPlanner.Create(job.SourceRange, job.TargetPath, workbookDirectory);
                planned.Add((plan, ReadAsStringGrid(plan.SourceRangeAddress)));
            }
        }
        catch (FormatException fex) { Warn(fex.Message); return; }
        catch (Exception ex)
        {
            Fail($"Export Wizard: failed to read a source range — {ex.Message}", ex);
            return;
        }

        // --- background thread: write each file --------------------------
        Task.Run(() =>
        {
            int written = 0;
            foreach (var (plan, rows) in planned)
            {
                int height = rows.GetLength(0);
                int width = rows.GetLength(1);
                if (height == 0 || width == 0)
                {
                    Warn($"Export Wizard: '{plan.SourceRangeAddress}' is empty — skipped.");
                    continue;
                }
                try
                {
                    EnsureDirectory(plan.AbsoluteTargetPath);
                    using var stream = new FileStream(
                        plan.AbsoluteTargetPath, FileMode.Create, FileAccess.Write, FileShare.Read);
                    CsvWriter.Write(
                        stream, EnumerateRows(rows),
                        delimiter: plan.Delimiter, lineTerminator: "\r\n",
                        encoding: null, writeBom: false);
                    written++;
                }
                catch (Exception ex)
                {
                    Fail($"Export Wizard: failed to write '{plan.AbsoluteTargetPath}' — {ex.Message}", ex);
                }
            }
            Trace.WriteLine($"Export Wizard: wrote {written}/{planned.Count} file(s).");
        });
    }

    /// <summary>Yield each row as an <see cref="System.Collections.Generic.IEnumerable{T}"/>
    /// of nullable strings so <see cref="CsvWriter.Write"/> can stream
    /// without materialising the whole CSV in memory. Iterator method
    /// keeps the grid alive via closure — safe because the caller
    /// retains it on the stack.</summary>
    private static System.Collections.Generic.IEnumerable<System.Collections.Generic.IEnumerable<string?>> EnumerateRows(string?[,] rows)
    {
        int height = rows.GetLength(0);
        int width = rows.GetLength(1);
        for (int r = 0; r < height; r++)
        {
            yield return RowSlice(rows, r, width);
        }
    }

    private static System.Collections.Generic.IEnumerable<string?> RowSlice(string?[,] rows, int r, int width)
    {
        for (int c = 0; c < width; c++)
            yield return rows[r, c];
    }

    /// <summary>Read a range into a stringified grid in one pass. Excel
    /// hands back <c>object[,]</c> for multi-cell ranges (1-based) or a
    /// scalar for a single cell; we normalise both into 0-based string
    /// grids ready for the CSV writer.</summary>
    private static string?[,] ReadAsStringGrid(string sourceAddress)
    {
        dynamic app = ExcelDnaUtil.Application;
        dynamic range = app.Range[sourceAddress];
        object? value = range.Value2;

        if (value is object[,] arr)
        {
            int r0 = arr.GetLowerBound(0);
            int c0 = arr.GetLowerBound(1);
            int height = arr.GetLength(0);
            int width = arr.GetLength(1);
            var grid = new string?[height, width];
            for (int i = 0; i < height; i++)
                for (int j = 0; j < width; j++)
                    grid[i, j] = FormatCell(arr[r0 + i, c0 + j]);
            return grid;
        }

        // Single cell.
        return new[,] { { FormatCell(value) } };
    }

    /// <summary>Format an Excel-side cell value (from
    /// <c>Range.Value2</c>) into the string form CSV expects. Strips
    /// Excel-DNA's <see cref="ExcelEmpty"/> / <see cref="ExcelMissing"/>
    /// / <see cref="ExcelError"/> sentinels (which can't live in
    /// <see cref="CsvCellFormatter"/> because that's cross-platform and
    /// the ExcelDna types are net48-only), then delegates the typed
    /// formatting to <see cref="CsvCellFormatter.Format"/>.</summary>
    private static string? FormatCell(object? value)
    {
        if (value is ExcelEmpty || value is ExcelMissing || value is ExcelError)
            return null;
        return CsvCellFormatter.Format(value);
    }

    /// <summary>Create the parent directory if it doesn't exist. A
    /// missing intermediate directory is a common user error (typing a
    /// fresh path that doesn't exist yet) and is cheap to fix here.</summary>
    private static void EnsureDirectory(string path)
    {
        var dir = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            Directory.CreateDirectory(dir!);
    }

    /// <summary>Workbook directory of the active workbook, or
    /// <see langword="null"/> for an unsaved workbook. See
    /// <see cref="ImportService"/> for the rationale on duplication.</summary>
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
            return null;
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
