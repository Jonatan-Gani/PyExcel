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
    // Last successfully-resolved identity. The ribbon's getEnabled callbacks
    // read these on every repaint, and the live Application.ActiveWorkbook COM
    // call can fail transiently whenever Excel is busy — a modal dialog is up,
    // a file dialog is open, focus is moving between windows. Returning null in
    // that window makes the whole ribbon grey out and (because nothing fires to
    // re-invalidate it) stay greyed. Caching the last good value and falling
    // back to it on a COM fault keeps the ribbon stable through those
    // transients. A genuinely-closed workbook is handled by the event sink,
    // which Forgets its state — so even a stale key resolves to a disabled
    // (Empty) state, which is the correct render.
    private volatile string? _cachedKey;
    private volatile string? _cachedDir;

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
                // wb == null can be a genuine "no workbook" or a transient
                // null while Excel is mid-operation; either way, keep serving
                // the last known key (a truly-closed book's state is Forgotten,
                // so it still renders disabled) rather than flickering to null.
                if (wb is null) return _cachedKey;

                // Defer the saved-vs-unsaved rule and the shared session
                // GUID to WorkbookKeys so the event sink agrees on keys.
                var key = WorkbookKeys.Resolve((string)wb.Name, (string)wb.Path, (string)wb.FullName);
                _cachedKey = key;
                return key;
            }
            catch
            {
                // COM busy / transient fault — keep the last known key so the
                // ribbon doesn't flicker to "disabled" mid-interaction.
                return _cachedKey;
            }
        }
    }

    public string? CurrentWorkbookDirectory
    {
        get
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                dynamic wb = app.ActiveWorkbook;
                if (wb is null) return _cachedDir;
                string path = (string)wb.Path;
                var dir = string.IsNullOrEmpty(path) ? null : path;
                _cachedDir = dir;
                return dir;
            }
            catch
            {
                return _cachedDir;
            }
        }
    }
}
