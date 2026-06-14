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
    private ScriptDirectoryWatcher? _scriptWatcher;
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

        // Let the ribbon kick a watcher re-sync after Enable provisions the
        // userScripts folder (that flow fires no WorkbookActivate).
        PyExcelServices.RequestScriptRefresh = SyncScriptWatcherForActive;
    }

    // -------------------------------------------------------------------------
    // Event handlers — every body wrapped by Guard
    // -------------------------------------------------------------------------

    private void OnWorkbookOpen(Excel.Workbook wb) => Guard(nameof(OnWorkbookOpen), () =>
    {
        RestoreWorkbookState(wb);
        SyncScriptWatcher(wb);
        InvalidateRibbon();
    });

    /// <summary>
    /// Restore the in-memory state for every workbook that is <em>already</em>
    /// open when the add-in loads. On a double-click / command-line launch
    /// Excel often opens the target workbook before the <c>.xll</c>'s
    /// <see cref="AddIn.AutoOpen"/> has subscribed this sink, so the
    /// <c>WorkbookOpen</c> event for that workbook is missed — and without this
    /// its saved "enabled" state and actions never load, so the user is wrongly
    /// asked to Enable again. Called once from <see cref="AddIn.AutoOpen"/>
    /// right after the sink is wired.
    /// </summary>
    public void RestoreOpenWorkbooks() => Guard(nameof(RestoreOpenWorkbooks), () =>
    {
        if (_app is null) return;
        foreach (Excel.Workbook wb in _app.Workbooks)
            RestoreWorkbookState(wb);

        // Point the script watcher at whatever's active now, then repaint so
        // the restored "enabled" state shows immediately.
        var active = _app.ActiveWorkbook;
        if (active is not null) SyncScriptWatcher(active);
        InvalidateRibbon();
    });

    /// <summary>Load <paramref name="wb"/>'s persisted v2 state (or migrate a
    /// v1 workbook's legacy Names) into the in-memory registry. Shared by
    /// <see cref="OnWorkbookOpen"/> and <see cref="RestoreOpenWorkbooks"/>.</summary>
    private void RestoreWorkbookState(Excel.Workbook wb)
    {
        string key = KeyOf(wb);
        var restored = WorkbookStatePersister.TryLoad(wb, key);
        if (restored is not null)
        {
            // Replace whatever transient Empty state existed with the
            // persisted one. Update validates the key matches.
            _state.Update(key, _ => restored);
            return;
        }

        // No v2 part — this may be a v1 workbook opened for the first time in
        // v2. Read its legacy defined Names and migrate them into a v2
        // CustomXMLPart so the user's saved actions/fields carry over.
        // Best-effort: a null reader result means there's nothing to migrate
        // (a brand-new or non-PyExcel workbook).
        var legacy = LegacyStateReader.TryRead(wb);
        if (legacy is not null)
        {
            var migrated = LegacyStateConverter.Convert(legacy, key);
            // Only populate the in-memory registry so the ribbon renders the
            // migrated state immediately. We deliberately don't write the v2
            // CustomXMLPart here — that would dirty the workbook just by
            // opening it. WorkbookBeforeSave flushes this state into the part
            // when the user actually saves, so the migration becomes durable then.
            _state.Update(key, _ => migrated);
        }
    }

    private void OnWorkbookActivate(Excel.Workbook wb) =>
        Guard(nameof(OnWorkbookActivate), () =>
        {
            // The active workbook changed: re-point the live script watcher at
            // its userScripts folder (which also repopulates AvailableScripts —
            // nothing else feeds it) and repaint.
            SyncScriptWatcher(wb);
            InvalidateRibbon();
        });

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

    /// <summary>The workbook's <c>userScripts</c> folder — the dedicated project
    /// folder chosen on Enable (saved in state) if set, else the workbook-derived
    /// default — the same rule the ribbon and KernelHost use, so all three agree
    /// on where scripts live. Null when there's no local folder (unsaved /
    /// cloud-URL workbook).</summary>
    private string? ResolveScriptsDir(Excel.Workbook wb)
    {
        var stored = _state.Get(KeyOf(wb)).ProjectDir;
        var workbookDir = string.IsNullOrEmpty(wb.Path) ? null : wb.Path;
        var projectDir = string.IsNullOrEmpty(stored)
            ? PyExcel.Common.ProjectDirectory.Resolve(workbookDir)
            : stored;
        return string.IsNullOrEmpty(projectDir)
            ? null
            : System.IO.Path.Combine(projectDir!, "userScripts");
    }

    /// <summary>(Re)point the live script watcher at the active workbook's
    /// userScripts folder so the ribbon's Script dropdown tracks files appearing
    /// / disappearing without a re-activate. The watcher pushes an initial
    /// snapshot from its constructor, so this also populates the list now. When
    /// the folder doesn't exist yet (workbook not Enabled) we drop the watcher
    /// and publish a one-shot snapshot instead — the list stays correct, and the
    /// watcher starts as soon as Enable creates the folder (via
    /// <see cref="PyExcelServices.RequestScriptRefresh"/>).</summary>
    private void SyncScriptWatcher(Excel.Workbook wb)
    {
        string key = KeyOf(wb);
        var scriptsDir = ResolveScriptsDir(wb);

        DisposeScriptWatcher();

        if (string.IsNullOrEmpty(scriptsDir) || !System.IO.Directory.Exists(scriptsDir))
        {
            _state.SetAvailableScripts(key, ScriptDirectoryWatcher.Snapshot(scriptsDir));
            return;
        }

        try
        {
            // The callback fires on the watcher's worker thread, but the
            // resulting StateChanged drives the ribbon's OnStateChanged, which
            // touches COM (CurrentWorkbookKey) — so marshal the state update onto
            // Excel's macro thread, per the watcher's documented contract.
            _scriptWatcher = new ScriptDirectoryWatcher(
                scriptsDir!,
                snapshot => ExcelAsyncUtil.QueueAsMacro(
                    () => _state.SetAvailableScripts(key, snapshot)));
        }
        catch (Exception ex)
        {
            // A watcher failure must never break activation — fall back to a scan.
            Trace.WriteLine($"AppEventSink.SyncScriptWatcher failed: {ex}");
            _state.SetAvailableScripts(key, ScriptDirectoryWatcher.Snapshot(scriptsDir));
        }
    }

    /// <summary>Re-sync the watcher for whatever workbook is active now. Wired to
    /// <see cref="PyExcelServices.RequestScriptRefresh"/> so the ribbon can start
    /// live watching the moment Enable provisions the userScripts folder.</summary>
    private void SyncScriptWatcherForActive() => Guard(nameof(SyncScriptWatcherForActive), () =>
    {
        var wb = _app?.ActiveWorkbook;
        if (wb is not null) SyncScriptWatcher(wb);
    });

    private void DisposeScriptWatcher()
    {
        var w = _scriptWatcher;
        _scriptWatcher = null;
        if (w is not null)
        {
            try { w.Dispose(); } catch { /* swallow */ }
        }
    }

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

        // Drop the hook and stop watching regardless of _app state. One sink
        // instance lives per add-in load (AutoOpen/AutoClose), so this is ours.
        PyExcelServices.RequestScriptRefresh = null;
        DisposeScriptWatcher();

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
