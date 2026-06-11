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
/// Drives the ribbon's Import button. Reads the user-configured source
/// file (CSV / TSV via <see cref="CsvParser"/>, or XLSX / XLSM / XLSB via
/// Excel COM) and writes its contents into the configured target range.
/// The cross-platform pieces — field validation, path resolution, format
/// detection, line parsing — live in <see cref="ImportPlanner"/> and
/// <see cref="CsvParser"/>; this class is the COM-bound shell that
/// dispatches them on the right threads.
///
/// <para><b>Threading / SAFE-1.</b> The ribbon callback returns
/// immediately. For the CSV path, file I/O happens on a background
/// <see cref="Task"/> and the write-back is queued onto Excel's main
/// thread via <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/>.
/// For the Excel-format path everything COM (open, read, close, write)
/// runs on the macro queue — opening a workbook via COM is itself a
/// main-thread operation, so there's no useful background work to split
/// out — but Excel's UI thread isn't blocked from the ribbon callback
/// itself.</para>
///
/// <para>Per-cell type inference for CSV matches Excel's built-in import:
/// a string that parses as a <see cref="double"/> under
/// <see cref="System.Globalization.CultureInfo.InvariantCulture"/>
/// becomes a numeric cell; <c>"TRUE"</c> / <c>"FALSE"</c>
/// (case-insensitive) become boolean cells; everything else stays a
/// string. Leading-zero strings (<c>"00123"</c>) intentionally fall
/// through to string-typed cells because parsing destroys the leading
/// zeros — the user's typing is the signal, the inference is a
/// convenience. Excel-format imports skip the inference step entirely —
/// the source workbook already carries cell types, which COM hands back
/// as native CLR values.</para>
/// </summary>
public static class ImportService
{
    /// <summary>
    /// Execute the import currently configured in <paramref name="state"/>.
    /// Returns immediately; the file read and the write-back happen
    /// asynchronously.
    /// </summary>
    /// <param name="state">The active workbook state carrying the Import
    /// source / target fields.</param>
    /// <param name="sheetChooser">Optional callback invoked (on the macro
    /// thread) when an Excel import needs the user to pick a sheet — the
    /// workbook has several and none was pinned with the <c>!Sheet</c>
    /// syntax. Returns the chosen sheet, or null to cancel the import.
    /// When null (no UI available), the import falls back to the first
    /// sheet, preserving the pre-picker behaviour.</param>
    public static void RunActiveImport(
        WorkbookState state,
        Func<IReadOnlyList<string>, string?>? sheetChooser = null)
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

        switch (plan.Format)
        {
            case ImportFormat.Csv:
                RunCsvImport(plan);
                break;
            case ImportFormat.Excel:
                RunExcelImport(plan, sheetChooser);
                break;
            default:
                Warn($"Import: unsupported format '{plan.Format}'.");
                break;
        }
    }

    // -------------------------------------------------------------------
    // CSV / TSV path
    // -------------------------------------------------------------------

    private static void RunCsvImport(ImportPlan plan)
    {
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
    private static object?[,] ToObjectGrid(
        System.Collections.Generic.IReadOnlyList<System.Collections.Generic.IReadOnlyList<string>> rows)
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

    // -------------------------------------------------------------------
    // Excel-format path (XLSX / XLSM / XLSB)
    // -------------------------------------------------------------------

    /// <summary>Dispatch the Excel-COM import. Pre-checks file existence
    /// on the main thread so the user gets a clean error without queuing
    /// an empty macro, then queues a single
    /// <see cref="ExcelAsyncUtil.QueueAsMacro(System.Action)"/> that
    /// opens, reads, closes, and writes — all COM, all on the macro
    /// queue (where the STA wants it).</summary>
    private static void RunExcelImport(
        ImportPlan plan,
        Func<IReadOnlyList<string>, string?>? sheetChooser)
    {
        if (!File.Exists(plan.AbsoluteSourcePath))
        {
            Warn($"Import: file not found at '{plan.AbsoluteSourcePath}'.");
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            object?[,]? table;
            try
            {
                table = ReadExcelSource(plan.AbsoluteSourcePath, plan.SheetName, sheetChooser);
            }
            catch (FormatException fex)
            {
                Warn(fex.Message);
                return;
            }
            catch (Exception ex)
            {
                Fail($"Import: failed to read '{plan.AbsoluteSourcePath}' — {ex.Message}", ex);
                return;
            }

            // Null = the user cancelled at the sheet picker — abort quietly.
            if (table is null)
            {
                Warn("Import: cancelled at sheet selection.");
                return;
            }

            if (table.GetLength(0) == 0 || table.GetLength(1) == 0)
            {
                Warn($"Import: '{plan.AbsoluteSourcePath}' parsed to zero rows.");
                return;
            }

            try { WriteTable(plan.TargetRangeAddress, table); }
            catch (Exception ex)
            {
                Fail(
                    $"Import: failed to write into '{plan.TargetRangeAddress}' — {ex.Message}",
                    ex);
            }
        });
    }

    /// <summary>Open the source workbook via COM, read the requested sheet's
    /// used range, and close it iff we were the one that opened it. If the
    /// workbook is already open in the running Excel instance (matching
    /// <c>FullName</c>), we reuse it and never close it — leaving the
    /// user's view alone.
    ///
    /// <para>Save / restore <c>ScreenUpdating</c> + <c>DisplayAlerts</c>
    /// around the call so a hidden Open doesn't flash the user's screen
    /// and a corrupt-file prompt is suppressed (and surfaced to the user
    /// via the normal exception path instead).</para>
    /// </summary>
    private static object?[,]? ReadExcelSource(
        string filePath,
        string? sheetName,
        Func<IReadOnlyList<string>, string?>? sheetChooser)
    {
        dynamic app = ExcelDnaUtil.Application;

        bool prevScreenUpdating = true;
        bool prevDisplayAlerts = true;
        try { prevScreenUpdating = (bool)app.ScreenUpdating; } catch { }
        try { prevDisplayAlerts = (bool)app.DisplayAlerts; } catch { }
        try { app.ScreenUpdating = false; } catch { }
        try { app.DisplayAlerts = false; } catch { }

        object? wbHandle = FindOpenWorkbook(app, filePath);
        bool weOpened = false;
        try
        {
            if (wbHandle is null)
            {
                // Workbooks.Open(Filename, UpdateLinks, ReadOnly, ...). The
                // remaining positional args default to Missing on the COM
                // side, which Excel treats as "use default for that arg".
                wbHandle = app.Workbooks.Open(filePath, 0, true);
                weOpened = true;
            }

            dynamic wb = wbHandle!;

            // Decide the sheet: a pinned !Sheet wins; otherwise one sheet
            // resolves automatically and several prompt the user (via the
            // injected chooser — null falls back to the first sheet).
            var resolution = SheetSelection.Resolve(sheetName, EnumerateSheetNames(wb));
            string? chosenSheet;
            switch (resolution.Kind)
            {
                case SheetResolutionKind.Empty:
                    throw new FormatException(
                        $"Import: '{filePath}' has no worksheets to import.");

                case SheetResolutionKind.Prompt:
                    if (sheetChooser is null)
                    {
                        // No UI available — preserve the pre-picker default.
                        chosenSheet = resolution.AvailableSheets[0];
                    }
                    else
                    {
                        chosenSheet = sheetChooser(resolution.AvailableSheets);
                        if (string.IsNullOrEmpty(chosenSheet))
                            return null; // cancelled
                    }
                    break;

                default: // Resolved
                    chosenSheet = resolution.Sheet;
                    break;
            }

            dynamic sheet;
            try
            {
                sheet = wb.Sheets[chosenSheet];
            }
            catch
            {
                throw new FormatException(
                    $"Import: sheet '{chosenSheet}' not found in '{filePath}'.");
            }

            dynamic used = sheet.UsedRange;
            object? value = used.Value2;
            return NormalizeUsedRange(value);
        }
        finally
        {
            if (wbHandle is not null && weOpened)
            {
                try { ((dynamic)wbHandle).Close(false); } catch { /* best-effort */ }
            }
            try { app.ScreenUpdating = prevScreenUpdating; } catch { }
            try { app.DisplayAlerts = prevDisplayAlerts; } catch { }
        }
    }

    /// <summary>List the workbook's worksheet names in tab order, so the
    /// sheet picker offers exactly what <c>Workbook.Sheets[name]</c> can
    /// later look up. Chart sheets are skipped — they have no
    /// <c>UsedRange</c> for the importer to read.</summary>
    private static IReadOnlyList<string> EnumerateSheetNames(dynamic wb)
    {
        var names = new List<string>();
        dynamic sheets = wb.Worksheets;
        int count = (int)sheets.Count;
        for (int i = 1; i <= count; i++)
        {
            try { names.Add((string)sheets[i].Name); }
            catch { /* skip an unreadable sheet rather than fail the import */ }
        }
        return names;
    }

    /// <summary>Walk the running app's open workbooks, returning the one
    /// whose <c>FullName</c> matches <paramref name="filePath"/>
    /// (case-insensitive). Returns <see langword="null"/> if no match —
    /// the caller then opens the file fresh.</summary>
    private static object? FindOpenWorkbook(dynamic app, string filePath)
    {
        try
        {
            dynamic wbs = app.Workbooks;
            int count = (int)wbs.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic wb = wbs[i];
                string fullName;
                try { fullName = (string)wb.FullName; }
                catch { continue; }
                if (string.Equals(fullName, filePath, StringComparison.OrdinalIgnoreCase))
                    return wb;
            }
        }
        catch
        {
            // Workbooks collection inaccessible — treat as no match.
        }
        return null;
    }

    /// <summary>Normalise the value of <c>UsedRange.Value2</c> into a
    /// 0-based <c>object?[,]</c> grid. Excel hands back a 1-based
    /// <c>object[,]</c> for multi-cell ranges, a scalar for a single
    /// cell, or <see langword="null"/> for an empty used range.</summary>
    private static object?[,] NormalizeUsedRange(object? value)
    {
        if (value is null)
            return new object?[0, 0];

        if (value is object[,] arr)
        {
            int r0 = arr.GetLowerBound(0);
            int c0 = arr.GetLowerBound(1);
            int height = arr.GetLength(0);
            int width = arr.GetLength(1);
            var grid = new object?[height, width];
            for (int i = 0; i < height; i++)
                for (int j = 0; j < width; j++)
                    grid[i, j] = NormalizeCell(arr[r0 + i, c0 + j]);
            return grid;
        }

        // Single cell — UsedRange of a workbook with one populated cell.
        return new[,] { { NormalizeCell(value) } };
    }

    /// <summary>Strip Excel-DNA's sentinel values from a cell hand-back.
    /// Errors map to <see langword="null"/> because the user is asking
    /// us to import the values, not the error chrome.</summary>
    private static object? NormalizeCell(object? value)
    {
        if (value is ExcelEmpty || value is ExcelMissing || value is ExcelError)
            return null;
        return value;
    }

    // -------------------------------------------------------------------
    // Shared helpers
    // -------------------------------------------------------------------

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
