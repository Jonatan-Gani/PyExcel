#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Drives the ribbon's Paste button: a plain paste of the OS clipboard into the
/// user-configured Destination range. Excel copies cells to the clipboard as
/// tab-separated text (with CSV-style quoting), so the clip is parsed with
/// <see cref="CsvParser.ParseTsv"/> and each field typed via
/// <see cref="CsvCellTypeInference"/> — numbers land as numbers, the rest as text
/// — exactly as the Import path treats a file.
///
/// <para><b>Independent of Python.</b> No venv, no kernel, no run archive: the only
/// inputs are the clipboard and the Destination address, so Paste works on any open
/// workbook whether or not it's been enabled for PyExcel.</para>
///
/// <para><b>Threading.</b> Must be called on Excel's main (STA) thread — the ribbon
/// callback already is. The clipboard read, the overwrite prompt, and the COM write
/// all run inline there; the payload is small and in-memory, so (unlike Import /
/// Export) there's no file I/O worth pushing onto a background task.</para>
///
/// <para><b>Overwrite confirmation.</b> Before writing, the service snapshots the
/// target range and asks <see cref="PastePreflight.RangeHasContent"/> whether the
/// paste would destroy data. If so, a <see cref="MessageBox"/> defaulting to "No"
/// must be confirmed; cancelling aborts without touching the sheet.</para>
/// </summary>
public static class PasteService
{
    /// <summary>
    /// Paste the clipboard into the Destination range configured in
    /// <paramref name="state"/> (<see cref="WorkbookState.PasteOutput"/>). Runs
    /// synchronously on the calling (main) thread.
    /// </summary>
    public static void RunActivePaste(WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));

        var raw = state.PasteOutput?.Trim();
        if (string.IsNullOrEmpty(raw))
        {
            Warn("Paste: no destination is set — click Edit to choose a target range " +
                 "(e.g. A1, or Sheet1!A1).");
            return;
        }
        string target = raw!;

        string? clip = ReadClipboardText();
        if (string.IsNullOrEmpty(clip))
        {
            Warn("Paste: the clipboard has no text to paste — copy some cells first.");
            return;
        }

        object?[,] grid;
        try
        {
            // Excel appends a trailing newline after the last row; drop it so we
            // don't paste a spurious empty row.
            var rows = CsvParser.ParseTsv(clip!.TrimEnd('\r', '\n'));
            grid = ToObjectGrid(rows);
        }
        catch (Exception ex)
        {
            Fail($"Paste: failed to parse the clipboard — {ex.Message}", ex);
            return;
        }

        int height = grid.GetLength(0);
        int width = grid.GetLength(1);
        if (height == 0 || width == 0)
        {
            Warn("Paste: the clipboard had nothing to paste.");
            return;
        }

        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic anchor = app.Range[target];
            dynamic targetRange = anchor.Resize[height, width];

            // Strip Excel-DNA's empty/missing sentinels so the cross-platform
            // preflight sees a clean snapshot, then prompt only when the paste
            // would overwrite real content.
            object? before = StripExcelSentinels(targetRange.Value2);
            if (PastePreflight.RangeHasContent(before)
                && !ConfirmOverwrite(target, height, width))
            {
                Trace.WriteLine($"Paste: user cancelled overwrite at '{target}'.");
                LogDisplay.WriteLine($"Paste: cancelled — '{target}' kept.");
                return;
            }

            targetRange.Value2 = grid;
            Trace.WriteLine(
                $"Paste: pasted {height}×{width} from the clipboard into '{target}'.");
        }
        catch (Exception ex)
        {
            Fail($"Paste: failed to write into '{target}' — {ex.Message}", ex);
        }
    }

    /// <summary>Read the clipboard's text (Excel's tab-separated cell copy, or any
    /// plain text). Returns null when the clipboard holds no text or can't be read —
    /// another process can briefly hold it locked, surfacing as an
    /// <see cref="System.Runtime.InteropServices.ExternalException"/>.</summary>
    private static string? ReadClipboardText()
    {
        try
        {
            return Clipboard.ContainsText() ? Clipboard.GetText() : null;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"Paste: clipboard read failed: {ex}");
            return null;
        }
    }

    /// <summary>Convert the parser's ragged record list into a rectangular
    /// <c>object?[,]</c>, typing each field (number / boolean / text) the same way
    /// the Import path does. Short rows are padded with nulls so every column
    /// lands.</summary>
    private static object?[,] ToObjectGrid(IReadOnlyList<IReadOnlyList<string>> rows)
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
                grid[r, c] = c < cols ? CsvCellTypeInference.Infer(row[c]) : null;
        }
        return grid;
    }

    /// <summary>Show the destructive-paste confirmation, defaulting to "No" so an
    /// accidental Enter can't overwrite data. Returns true iff the user clicked
    /// Yes.</summary>
    private static bool ConfirmOverwrite(string targetAddress, int rows, int cols)
    {
        var prompt =
            $"The target range '{targetAddress}' ({rows}×{cols}) already contains " +
            "values that the paste will overwrite." +
            Environment.NewLine + Environment.NewLine + "Continue?";

        var answer = MessageBox.Show(
            prompt,
            "PyExcel — confirm overwrite",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning,
            MessageBoxDefaultButton.Button2);
        return answer == DialogResult.Yes;
    }

    /// <summary>Replace Excel-DNA's <see cref="ExcelEmpty"/> / <see cref="ExcelMissing"/>
    /// sentinels with null so the cross-platform <see cref="PastePreflight"/> sees a
    /// clean snapshot. <see cref="ExcelError"/> is preserved — an error cell is still
    /// content a paste would destroy.</summary>
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
                    cleaned[i, j] = (cell is ExcelEmpty || cell is ExcelMissing) ? null : cell;
                }
            return cleaned;
        }

        return value2;
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
