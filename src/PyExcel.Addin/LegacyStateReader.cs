#if NETFRAMEWORK
using System;
using System.Diagnostics;
using PyExcel.State;

namespace PyExcel.Addin;

// Declared inside the namespace deliberately — see the note in AppEventSink.cs.
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// The Windows-only COM half of the v1 → v2 state migration: reads the v1
/// PyExcel defined Names off a live workbook into a
/// <see cref="LegacyWorkbookState"/> that the cross-platform
/// <see cref="LegacyStateConverter"/> can turn into a v2
/// <see cref="WorkbookState"/>.
///
/// <para>v1 stored its state per <em>sheet</em> (each as a sheet-scoped Name)
/// plus the workbook-scoped <c>PyExcelEnabled</c> flag; v2 holds one
/// workbook-scoped state. With no lossless multi-sheet collapse, this reader
/// carries forward the <em>first</em> worksheet that has any PyExcel state
/// (<see cref="LegacyStateConverter.HasSheetContent"/>). If no sheet has state
/// but the workbook was enabled, an enabled-only legacy record is returned so
/// the toggle survives; otherwise <see langword="null"/> (nothing to
/// migrate).</para>
///
/// <para>Values are recovered from each Name's <c>RefersTo</c> formula via the
/// pure <see cref="LegacyFormulaDecoder"/> — no <c>Evaluate</c>, so reading
/// neither forces a recalc nor depends on the name's scope being active. The
/// whole read is best-effort: a COM fault on any Name is swallowed (that Name
/// reads as absent) so a malformed workbook can never abort the open hook.</para>
/// </summary>
internal static class LegacyStateReader
{
    /// <summary>Read the v1 state to migrate from <paramref name="workbook"/>,
    /// or <see langword="null"/> if the workbook carries no migratable PyExcel
    /// state.</summary>
    public static LegacyWorkbookState? TryRead(Excel.Workbook workbook)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        try
        {
            string? enabled = ReadFromNames(workbook.Names, LegacyStateConverter.LegacyNames.Enabled);

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                var legacy = ReadSheet(ws, enabled);
                if (LegacyStateConverter.HasSheetContent(legacy))
                    return legacy;
            }

            // No sheet carried per-sheet state. Preserve the enabled toggle
            // alone if it was set; otherwise there is nothing to migrate.
            return IsTruthy(enabled)
                ? new LegacyWorkbookState { Enabled = enabled }
                : null;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"LegacyStateReader.TryRead failed: {ex}");
            return null;
        }
    }

    private static LegacyWorkbookState ReadSheet(Excel.Worksheet ws, string? enabled)
    {
        Excel.Names names = ws.Names;
        var N = LegacyStateConverter.LegacyNames;
        return new LegacyWorkbookState
        {
            Enabled = enabled,
            SelectedAction = ReadFromNames(names, N.SelectedAction),
            Actions = ReadFromNames(names, N.Actions),
            SelectedScript = ReadFromNames(names, N.SelectedScript),
            PyInput = ReadFromNames(names, N.PyInput),
            PyOutput = ReadFromNames(names, N.PyOutput),
            ImportInput = ReadFromNames(names, N.ImportInput),
            ImportOutput = ReadFromNames(names, N.ImportOutput),
            ExportInput = ReadFromNames(names, N.ExportInput),
            ExportOutput = ReadFromNames(names, N.ExportOutput),
            PasteOutput = ReadFromNames(names, N.PasteOutput),
        };
    }

    /// <summary>Look up one Name in <paramref name="names"/> and decode its
    /// <c>RefersTo</c> formula. Returns <see langword="null"/> when the Name is
    /// absent (the indexer throws) or its formula isn't a v1 string value.</summary>
    private static string? ReadFromNames(Excel.Names names, string id)
    {
        try
        {
            Excel.Name nm = names[id];
            return LegacyFormulaDecoder.Decode(nm.RefersTo as string);
        }
        catch
        {
            // Name not defined on this scope (or unreadable) — treat as absent.
            return null;
        }
    }

    private static bool IsTruthy(string? value)
    {
        var v = value?.Trim();
        return string.Equals(v, "1", StringComparison.Ordinal)
            || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
    }
}
#endif
