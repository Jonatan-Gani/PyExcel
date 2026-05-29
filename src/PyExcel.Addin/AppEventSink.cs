#if NETFRAMEWORK
using System;
using System.Diagnostics;
using ExcelDna.Integration;
using PyExcel.State;

namespace PyExcel.Addin;

// The Excel alias MUST be declared here, inside the namespace, rather than at
// the top of the file. At file scope the alias lands in the global namespace,
// so the bare name `Excel` resolves outward from PyExcel.Addin, reaches the
// enclosing PyExcel namespace, and binds to our own PyExcel.Excel before the
// alias is ever consulted — producing CS0234 on Excel.Application /
// Excel.Workbook. Declared inside PyExcel.Addin the alias is found first and
// correctly maps `Excel` to the embedded Office PIA.
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Subscribes to the Excel <see cref="Excel.Application"/> events that
/// per-workbook state depends on, and keeps the <see cref="StateService"/>
/// and the ribbon in step with what the user is doing:
///
/// <list type="bullet">
///   <item><c>WorkbookOpen</c> → restore saved state from the workbook's
///     CustomXMLPart (via <see cref="WorkbookStatePersister"/>) and repaint
///     the ribbon.</item>
///   <item><c>WorkbookActivate</c> → repaint the ribbon: the active
///     workbook key changed, so every getter now renders a different
///     state, but nothing in the registry mutated so no
///     <see cref="StateService.StateChanged"/> fired on its own.</item>
///   <item><c>WorkbookBeforeSave</c> → flush the current state into the
///     workbook's CustomXMLPart so it travels with the file.</item>
///   <item><c>WorkbookBeforeClose</c> → flush state, then
///     <see cref="StateService.Forget"/> the key so closed workbooks don't
///     accumulate in memory.</item>
///   <item><c>SheetActivate</c> → record the active sheet name.</item>
/// </list>
///
/// <para><b>Typed PIA.</b> Application events can't be wired late-bound
/// through <c>dynamic</c> — there is no <c>+=</c> over a dispatch object —
/// so this is the one place PyExcel takes a typed Office interop reference.
/// It comes from the <c>ExcelDna.Interop</c> NuGet package, whose build
/// <c>.targets</c> add the assembly references, so a CI runner with no
/// Excel installed still compiles this. Everything else COM-bound in
/// PyExcel (<see cref="ExcelWorkbookContext"/>, the range runner) stays
/// late-bound.</para>
///
/// <para><b>Safety.</b> Every handler body runs inside <see cref="Guard"/>
/// so a fault can never propagate back into Excel's event pump (which can
/// destabilise the host). Lifetime is the add-in lifetime:
/// <see cref="AddIn.AutoOpen"/> constructs the sink,
/// <see cref="AddIn.AutoClose"/> disposes it, and <see cref="Dispose"/>
/// unsubscribes every handler.</para>
/// </summary>
internal sealed class AppEventSink : IDisposable
{
    private readonly StateService _state;
    private readonly IWorkbookContext _context;
    private Excel.Application? _app;
    private bool _disposed;

    /// <summary>Subscribe to the Application events. Must be called on
    /// Excel's main (STA) thread — <see cref="AddIn.AutoOpen"/> satisfies
    /// this — because <c>IConnectionPoint.Advise</c> (which the typed
    /// <c>+=</c> performs under the hood) requires an STA caller.</summary>
    public AppEventSink(StateService state, IWorkbookContext context)
    {
        _state = state ?? throw new ArgumentNullException(nameof(state));
        _context = context ?? throw new ArgumentNullException(nameof(context));

        // ExcelDnaUtil.Application is the live Excel Application RCW; the
        // cast succeeds because the embedded interop interface is just a
        // typed view onto the same COM object.
        _app = (Excel.Application)ExcelDnaUtil.Application;
        _app.WorkbookOpen += OnWorkbookOpen;
        _app.WorkbookActivate += OnWorkbookActivate;
        _app.WorkbookBeforeSave += OnWorkbookBeforeSave;
        _app.WorkbookBeforeClose += OnWorkbookBeforeClose;
        _app.SheetActivate += OnSheetActivate;
    }

    // -------------------------------------------------------------------------
    // Event handlers — every body wrapped by Guard
    // -------------------------------------------------------------------------

    private void OnWorkbookOpen(Excel.Workbook wb) => Guard(nameof(OnWorkbookOpen), () =>
    {
        string key = KeyOf(wb);
        var restored = WorkbookStatePersister.TryLoad(wb, key);
        if (restored is not null)
        {
            // Replace whatever transient Empty state existed with the
            // persisted one. Update validates the key matches.
            _state.Update(key, _ => restored);
        }
        InvalidateRibbon();
    });

    private void OnWorkbookActivate(Excel.Workbook wb) =>
        Guard(nameof(OnWorkbookActivate), InvalidateRibbon);

    private void OnWorkbookBeforeSave(Excel.Workbook wb, bool saveAsUI, ref bool cancel) =>
        Guard(nameof(OnWorkbookBeforeSave), () =>
        {
            string key = KeyOf(wb);
            WorkbookStatePersister.Save(wb, _state.Get(key));
        });

    private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel) =>
        Guard(nameof(OnWorkbookBeforeClose), () =>
        {
            // Persist before forgetting so the part is current even if the
            // user wasn't prompted to save (the close itself doesn't write
            // the file, but a later reopen of an already-saved workbook
            // should still see fresh state). If the user cancels the close,
            // the part stays intact and re-activating re-reads from memory —
            // which is why we don't touch the file here, only the part and
            // the in-memory registry.
            string key = KeyOf(wb);
            WorkbookStatePersister.Save(wb, _state.Get(key));
            _state.Forget(key);
        });

    private void OnSheetActivate(object sh) => Guard(nameof(OnSheetActivate), () =>
    {
        var key = _context.CurrentWorkbookKey;
        if (key is null) return;
        // Worksheet and Chart both expose Name; read it late-bound so we
        // don't have to know which kind of sheet just activated.
        dynamic sheet = sh;
        _state.SetCurrentSheet(key, (string)sheet.Name);
    });

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static string KeyOf(Excel.Workbook wb) =>
        WorkbookKeys.Resolve(wb.Name, wb.Path, wb.FullName);

    private static void InvalidateRibbon()
    {
        // The ribbon registers this in RibbonOnLoad; it queues
        // IRibbonUI.Invalidate onto the macro thread. Null before the
        // ribbon loads / after unload — a no-op then.
        PyExcelServices.RequestRibbonInvalidate?.Invoke();
    }

    private static void Guard(string handler, Action body)
    {
        try
        {
            body();
        }
        catch (Exception ex)
        {
            // Never let an exception cross back into Excel's event pump.
            Trace.WriteLine($"AppEventSink.{handler} failed: {ex}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_app is null) return;
        try
        {
            _app.WorkbookOpen -= OnWorkbookOpen;
            _app.WorkbookActivate -= OnWorkbookActivate;
            _app.WorkbookBeforeSave -= OnWorkbookBeforeSave;
            _app.WorkbookBeforeClose -= OnWorkbookBeforeClose;
            _app.SheetActivate -= OnSheetActivate;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"AppEventSink.Dispose failed to unsubscribe: {ex}");
        }
        finally
        {
            _app = null;
        }
    }
}
#endif
