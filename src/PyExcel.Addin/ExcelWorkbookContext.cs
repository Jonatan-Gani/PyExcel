using ExcelDna.Integration;
using PyExcel.State;

namespace PyExcel.Addin;

/// <summary>
/// Concrete <see cref="IWorkbookContext"/> backed by
/// <c>Application.ActiveWorkbook</c> over the Excel COM interop. Wired
/// into <see cref="PyExcelServices.WorkbookContext"/> by
/// <see cref="AddIn.AutoOpen"/>; tests use a hand-written fake instead
/// of this class so the cross-platform CI slice doesn't need Excel.
///
/// <para>Key strategy:</para>
/// <list type="bullet">
///   <item>Saved workbooks → <c>Workbook.FullName</c> (the full path
///     including the filename). Stable across closes/reopens because
///     the path is the on-disk identity.</item>
///   <item>Unsaved workbooks (no on-disk path) → a synthetic
///     <c>"unsaved:{SessionGuid}:{Workbook.Name}"</c> key. The
///     saved-vs-unsaved decision and the per-add-in-load session GUID
///     both live in <see cref="WorkbookKeys"/> so the COM event sink
///     derives identical keys for the same workbook — a "Save As"
///     promotes the workbook to a path-based key on the next access
///     (the previous unsaved key is then orphaned in the
///     <see cref="StateService"/> until <c>WorkbookBeforeClose</c>
///     calls <see cref="StateService.Forget"/>).</item>
/// </list>
///
/// <para>Returns <see langword="null"/> when Excel has no active
/// workbook — typically during add-in load before any workbook is
/// open. The ribbon renders as "all disabled" in that state, which is
/// what we want.</para>
/// </summary>
public sealed class ExcelWorkbookContext : IWorkbookContext
{
    public string? CurrentWorkbookKey
    {
        get
        {
            try
            {
                // dynamic instead of the typed COM PIA: keeps the active-
                // workbook lookup free of an Office interop reference,
                // matching the Excel-DNA convention. (AppEventSink is the
                // one place a typed reference is unavoidable — events can't
                // be wired late-bound.)
                dynamic app = ExcelDnaUtil.Application;
                dynamic wb = app.ActiveWorkbook;
                if (wb is null) return null;

                // Defer the saved-vs-unsaved rule and the shared session
                // GUID to WorkbookKeys so the event sink agrees on keys.
                return WorkbookKeys.Resolve((string)wb.Name, (string)wb.Path, (string)wb.FullName);
            }
            catch
            {
                // COM exceptions during shutdown / between workbook events
                // can leave ActiveWorkbook in a transient state. The
                // ribbon's getters all tolerate a null key (they render as
                // "no workbook"), so a swallow-and-return-null here is the
                // right answer rather than crashing the ribbon callback.
                return null;
            }
        }
    }
}
