using System;
using ExcelDna.Integration;
using PyExcel.Common.Logging;
using PyExcel.State;

namespace PyExcel.Addin;

/// <summary>
/// Excel-DNA add-in entry point. Lifetime matches the .xll: Excel loads us
/// once on startup, unloads us once on quit. This class is therefore the
/// closest thing v2 has to the v1 <c>Workbook_Open</c> handler — except it
/// fires before any user workbook is active, which is what we want.
/// </summary>
/// <remarks>
/// What this method does NOT do, by deliberate design (see plan §SAFE-2):
/// it does not spawn the Python kernel, it does not enumerate user scripts,
/// it does not touch any per-workbook state. The kernel is started lazily
/// on first <c>OnRunPython</c>; user-script discovery is event-driven via
/// <see cref="System.IO.FileSystemWatcher"/>. <c>AutoOpen</c> exists only
/// to bring the singleton services online — the logger and (in later
/// phases) the AppEventSink and KernelSupervisor.
/// </remarks>
public sealed class AddIn : IExcelAddIn
{
    // Keep service singletons rooted on the AddIn instance so they survive
    // until AutoClose. Excel-DNA holds the AddIn itself, which keeps these
    // alive — no static lifetime tricks needed.
    private ILog? _log;
    private AppEventSink? _appEventSink;

    public void AutoOpen()
    {
        try
        {
            _log = new FileLog();
            _log.Info($"PyExcel v{ThisAssemblyVersion()} AutoOpen");

            // Bitness check per plan risk #4: refuse to load on 32-bit Excel.
            if (IntPtr.Size == 4)
            {
                _log.Error("PyExcel v2 requires 64-bit Excel; running 32-bit detected.");
                ExcelDna.Integration.XlCall.Excel(
                    XlCall.xlcAlert,
                    "PyExcel v2 requires 64-bit Excel.\n\n" +
                    "Please install the 64-bit version of Microsoft Excel to use this add-in.",
                    2 /* xlAlertWarning */);
                return;
            }

            // Phase 3: wire the concrete WorkbookContext into the service
            // locator the ribbon callbacks read from. The default is a
            // NullWorkbookContext that always returns "no active workbook";
            // swapping it here means the ribbon starts reflecting real
            // workbook identity as soon as AutoOpen finishes.
            PyExcelServices.WorkbookContext = new ExcelWorkbookContext();

            // Phase 3: subscribe to Application events so per-workbook state
            // stays in step with open / activate / save / close / sheet-change,
            // and the ribbon repaints when the active workbook changes.
            // Degrade gracefully — without the sink the ribbon and =PY.RUN
            // still work, you just lose auto-refresh-on-activate and state
            // persistence, so a sink failure must not fail the whole AutoOpen.
            try
            {
                _appEventSink = new AppEventSink(
                    PyExcelServices.State, PyExcelServices.WorkbookContext);
                _log.Info("AppEventSink subscribed to Application events");

                // Restore state for any workbook already open at load time. On a
                // double-click / command-line launch Excel often opens the target
                // workbook before this add-in finishes loading, so its WorkbookOpen
                // event is missed — without this, a saved-and-enabled workbook
                // comes up looking un-enabled and the user is asked to Enable again.
                _appEventSink.RestoreOpenWorkbooks();
            }
            catch (Exception ex)
            {
                _log.Error("AppEventSink wiring failed; continuing without it", ex);
            }

            // Subsequent phases bring more services online here:
            //   Phase 2: _kernelSupervisor = new KernelSupervisor(_log);
        }
        catch (Exception ex)
        {
            // We must not let an exception escape AutoOpen — Excel will
            // disable the add-in and leave the user with no UI affordance
            // to re-enable it. Log and degrade.
            _log ??= NullLog.Instance;
            _log.Error("AutoOpen failed", ex);
        }
    }

    public void AutoClose()
    {
        try
        {
            _log?.Info("PyExcel AutoClose");
            // Tear down in reverse-init order.
            //   Phase 2 (when added): _kernelSupervisor?.DrainAndDispose(timeoutMs: 3000);
            _appEventSink?.Dispose();
            _appEventSink = null;
            // Drop the ribbon-invalidate hook so a stale closure can't fire
            // against a torn-down ribbon after unload.
            PyExcelServices.RequestRibbonInvalidate = null;
        }
        catch (Exception ex)
        {
            _log?.Error("AutoClose failed", ex);
        }
    }

    private static string ThisAssemblyVersion()
    {
        var name = typeof(AddIn).Assembly.GetName();
        return name.Version?.ToString() ?? "unknown";
    }
}
