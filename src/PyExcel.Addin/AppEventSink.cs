#if NETFRAMEWORK
using System;
using System.Diagnostics;
using ExcelDna.Integration;
using PyExcel.Common.Logging;
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
///     CustomXMLPart (via <see cref="WorkbookStatePersister"/>), validate the
///     on-disk structure to set the readiness gate, load the live elements,
///     and repaint the ribbon.</item>
///   <item><c>WorkbookActivate</c> → bring the now-active workbook online
///     (restore-if-empty, re-validate its structure, reload elements) and
///     repaint: the active workbook key changed, so every getter now renders a
///     different state, but nothing in the registry mutated so no
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
    private readonly ILog _log = new FileLog();
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
        _app.WorkbookAfterSave += OnWorkbookAfterSave;
        _app.WorkbookBeforeClose += OnWorkbookBeforeClose;
        _app.SheetActivate += OnSheetActivate;

        // Let the ribbon kick a watcher re-sync after Enable provisions the
        // userScripts folder (that flow fires no WorkbookActivate).
        PyExcelServices.RequestScriptRefresh = SyncScriptWatcherForActive;

        _log.Info("AppEventSink: constructed and subscribed to Application events");
    }

    // -------------------------------------------------------------------------
    // Event handlers — every body wrapped by Guard
    // -------------------------------------------------------------------------

    private void OnWorkbookOpen(Excel.Workbook wb) => Guard(nameof(OnWorkbookOpen), () =>
    {
        // Restore + validate, then load the live elements and repaint. Alert on a
        // broken environment so the user learns on open, not at Run.
        BringOnline(wb, alertOnIncomplete: true);
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
        var app = _app;
        if (app is null) return;
        foreach (Excel.Workbook wb in app.Workbooks)
            BringOnline(wb, alertOnIncomplete: true);

        // Load the live elements for whatever's active now, then repaint so the
        // restored, validated state shows immediately.
        var active = app.ActiveWorkbook;
        if (active is not null) SyncScriptWatcher(active);
        InvalidateRibbon();
    });

    /// <summary>
    /// Ensure <paramref name="wb"/>'s saved PyExcel state is loaded into the
    /// in-memory registry — the add-in asking "is this workbook already a
    /// PyExcel project?" every time it sees one (open, activate, or load-time
    /// scan). Load order: the profile embedded in the workbook
    /// (<see cref="WorkbookStatePersister"/>, the single source of truth), then a
    /// one-time migration from an external sidecar an earlier build wrote, then a
    /// v1 (defined-Names) migration.
    ///
    /// <para><b>Load-if-empty.</b> We only load when the in-memory state has
    /// nothing meaningful yet, so re-activating a workbook the user is actively
    /// editing never clobbers unsaved in-memory changes — but a freshly opened
    /// (or just-closed-and-reopened, hence Forgotten) workbook always gets its
    /// project restored.</para>
    /// </summary>
    private void EnsureRestored(Excel.Workbook wb)
    {
        string key = KeyOf(wb);
        if (_state.GetProfile(key).IsMeaningful)
        {
            _log.Info($"EnsureRestored: '{key}' already has state in memory; skip");
            return;
        }

        // Single source of truth: the per-sheet profile embedded in the workbook.
        // Fallback: a one-time, read-only migration from the external sidecar an
        // earlier build wrote — once read it persists into the in-workbook part on
        // the next save (and the save hook then retires the sidecar).
        var restored = WorkbookStatePersister.TryLoad(wb, key)
                       ?? MigrateLegacySidecar(wb, key);
        if (restored is not null)
        {
            _state.LoadProfile(key, restored);
            _log.Info($"EnsureRestored: key='{key}' restored (sheets={restored.Sheets.Count})");
            return;
        }

        // No v2 state anywhere — this may be a v1 workbook opened for the first
        // time in v2. Read its legacy defined Names and migrate the one carried
        // sheet into the per-sheet model's default bucket. Best-effort: a null
        // reader result means there's nothing to migrate. We don't write the
        // CustomXMLPart here — that would dirty the workbook just by opening it; a
        // later save flushes it.
        var legacy = LegacyStateReader.TryRead(wb);
        if (legacy is not null)
            _state.LoadProfile(key, WorkbookProfileData.FromState(LegacyStateConverter.Convert(legacy, key)));
    }

    private void OnWorkbookActivate(Excel.Workbook wb) =>
        Guard(nameof(OnWorkbookActivate), () =>
        {
            // The active workbook changed. Bring it online (restore if-empty so we
            // never clobber a workbook being edited; point the current-sheet pointer
            // at its live active sheet; re-validate its structure), then re-point the
            // live script watcher at its userScripts folder and repaint. Re-validating
            // here — cheap, file-only — is what makes the gate smart: a structure that
            // broke (or was repaired) while another workbook was foreground is picked
            // up the moment the user switches back. No nag alert on activate, though;
            // that fires once on open.
            BringOnline(wb, alertOnIncomplete: false);
            SyncScriptWatcher(wb);
            InvalidateRibbon();
        });

    private void OnWorkbookBeforeSave(Excel.Workbook wb, bool saveAsUI, ref bool cancel) =>
        Guard(nameof(OnWorkbookBeforeSave), () =>
        {
            string key = KeyOf(wb);
            var profile = _state.GetProfile(key);
            // Only a PyExcel project gets a profile part. Never tattoo a plain
            // workbook with an empty part just because the user saved it — a
            // workbook becomes a project via Enable, which triggers its own save.
            if (!profile.IsMeaningful)
            {
                _log.Info($"OnWorkbookBeforeSave: key='{key}' not a PyExcel project; skip");
                return;
            }
            var projectDir = ResolveProjectDir(wb);
            _log.Info($"OnWorkbookBeforeSave: key='{key}' dir='{projectDir}' " +
                      $"enabled={profile.Enabled} sheets={profile.Sheets.Count}");
            // The workbook is the single store: flush the profile into its
            // embedded CustomXMLPart so it lands in the file as part of this save.
            WorkbookStatePersister.Save(wb, profile, projectDir, wb.Name, NullIfEmpty(wb.FullName));
        });

    private void OnWorkbookAfterSave(Excel.Workbook wb, bool success) =>
        Guard(nameof(OnWorkbookAfterSave), () =>
        {
            // The embedded profile part is now on disk (written in BeforeSave).
            // Only now is it safe to retire any external profile sidecar a
            // previous build left — never before the replacement is persisted, so
            // a failed save can't lose state.
            if (success) DeleteLegacySidecars(wb);
        });

    private void OnWorkbookBeforeClose(Excel.Workbook wb, ref bool cancel) =>
        Guard(nameof(OnWorkbookBeforeClose), () =>
        {
            string key = KeyOf(wb);
            var profile = _state.GetProfile(key);
            // Flush the profile only when there's something to flush AND the
            // workbook is already dirty. Writing a CustomXMLPart marks the
            // workbook changed, so flushing a clean workbook would pop a spurious
            // "save changes?" prompt on close — and a clean workbook's last save
            // already wrote the current state, so there's nothing to flush. When
            // it IS dirty, writing now means a Save from the close prompt (or a
            // cancelled close) keeps the latest in-memory state.
            bool dirty = IsDirty(wb);
            var projectDir = ResolveProjectDir(wb);
            _log.Info($"OnWorkbookBeforeClose: key='{key}' dir='{projectDir}' " +
                      $"enabled={profile.Enabled} dirty={dirty}");
            if (dirty && profile.IsMeaningful)
                WorkbookStatePersister.Save(wb, profile, projectDir, wb.Name, NullIfEmpty(wb.FullName));
            _state.Forget(key);
            PyExcelServices.Health.Clear(key);
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
        var projectDir = ResolveProjectDir(wb);
        return string.IsNullOrEmpty(projectDir)
            ? null
            : System.IO.Path.Combine(projectDir!, "userScripts");
    }

    /// <summary>The workbook's project folder: the dedicated folder chosen on
    /// Enable (saved in state) if set, else the workbook-derived default — the
    /// same rule the ribbon and KernelHost use. This is where the project
    /// profile (<c>pyexcel.project.xml</c>), the venv, the kernel and
    /// userScripts all live. Null for an unsaved / cloud-without-local
    /// workbook.</summary>
    private string? ResolveProjectDir(Excel.Workbook wb)
    {
        var stored = _state.Get(KeyOf(wb)).ProjectDir;
        var workbookDir = string.IsNullOrEmpty(wb.Path) ? null : wb.Path;
        return string.IsNullOrEmpty(stored)
            ? PyExcel.Common.ProjectDirectory.Resolve(workbookDir)
            : stored;
    }

    private static string? NullIfEmpty(string? s) => string.IsNullOrEmpty(s) ? null : s;

    /// <summary>Whether the workbook has unsaved changes (<c>Workbook.Saved</c> is
    /// false). On a COM hiccup, assume clean so we never force a spurious save
    /// prompt by flushing the profile part into an otherwise-untouched workbook.</summary>
    private static bool IsDirty(Excel.Workbook wb)
    {
        try { return !wb.Saved; }
        catch { return false; }
    }

    /// <summary>One-time, read-only migration from the external profile sidecar an
    /// earlier build wrote (<c>&lt;projectDir&gt;/.pyexcel/project.xml</c>, or the
    /// even older loose <c>&lt;projectDir&gt;/pyexcel.project.xml</c>) into the
    /// in-workbook part. The add-in no longer writes a sidecar — the workbook
    /// itself is the single store — so this only carries forward state from a
    /// workbook enabled by a prior build. Returns the migrated state, or
    /// <see langword="null"/> if there's no readable sidecar. The file is left in
    /// place until a save flushes the in-workbook part, after which
    /// <see cref="OnWorkbookAfterSave"/> removes it — so state is never deleted
    /// before its replacement is on disk.</summary>
    private WorkbookProfileData? MigrateLegacySidecar(Excel.Workbook wb, string key)
    {
        foreach (var path in LegacySidecarPaths(wb))
        {
            try
            {
                if (!System.IO.File.Exists(path)) continue;
                if (ProjectProfileCodec.TryDeserialize(
                        System.IO.File.ReadAllText(path), key, out var data, out _)
                    && data is not null)
                {
                    _log.Info($"MigrateLegacySidecar: migrated '{path}' for '{key}'");
                    return data;
                }
            }
            catch (Exception ex)
            {
                _log.Error($"MigrateLegacySidecar: failed reading '{path}'", ex);
            }
        }
        return null;
    }

    /// <summary>Bring a workbook "online" in the in-memory model — the consolidated
    /// open / activate / load-time path. In one place it: restores the workbook's
    /// saved PyExcel state (load-if-empty, so it never clobbers unsaved edits), points
    /// the current-sheet pointer at its live active sheet, then refreshes the
    /// readiness verdict the ribbon gates on. Loading the live <em>elements</em> (the
    /// script watcher) is the caller's next step, done once for the foreground
    /// workbook — and only when the project is Ready, so validation and loading move
    /// together.</summary>
    private void BringOnline(Excel.Workbook wb, bool alertOnIncomplete)
    {
        EnsureRestored(wb);
        SetCurrentSheetToLive(wb);
        RefreshReadiness(wb, alertOnIncomplete);
    }

    /// <summary>For an <em>enabled</em> workbook, validate that its on-disk project
    /// structure (venv, kernel, userScripts) is present and record the verdict in
    /// <see cref="PyExcelServices.Health"/> — the single gate the ribbon turns into
    /// the enabled-state of both the data controls and the Enable/Repair button, so
    /// the controls go live (and the elements load) only once the environment is
    /// confirmed whole. Cheap (file-existence only), so it's safe to run on every open
    /// and activate. A non-enabled workbook clears any stale result. When the
    /// structure is incomplete and <paramref name="alertOnIncomplete"/> is set,
    /// surface a one-time modal so the user learns on open — not at Run — that the
    /// environment needs repair.</summary>
    private void RefreshReadiness(Excel.Workbook wb, bool alertOnIncomplete)
    {
        string key = KeyOf(wb);
        if (!_state.GetProfile(key).Enabled)
        {
            PyExcelServices.Health.Clear(key);
            return;
        }
        var projectDir = ResolveProjectDir(wb);
        var check = ProjectStructureValidator.Validate(projectDir);
        PyExcelServices.Health.Set(key, check);
        _log.Info($"RefreshReadiness: key='{key}' dir='{projectDir}' ok={check.Ok} " +
                  $"missing='{string.Join(", ", check.Missing)}'");
        if (!check.Ok && alertOnIncomplete) QueueStructureAlert(wb, check);
    }

    /// <summary>Show a modal alert (deferred onto the macro queue so it doesn't run
    /// inside the open event pump) telling the user the workbook's PyExcel
    /// environment is incomplete and how to repair it.</summary>
    private void QueueStructureAlert(Excel.Workbook wb, ProjectStructureCheck check)
    {
        string name;
        try { name = wb.Name; } catch { name = "this workbook"; }

        var body = new System.Text.StringBuilder();
        body.Append("PyExcel — environment incomplete\n\n");
        body.Append('"').Append(name).Append("\" is set up for PyExcel, but its ");
        body.Append("Python environment is missing or incomplete:\n");
        foreach (var m in check.Missing) body.Append("  • ").Append(m).Append('\n');
        body.Append("\nOpen the PyExcel tab and click Enable to reinstall it. ");
        body.Append("Until then, Run won't work.");
        var message = body.ToString();

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try { XlCall.Excel(XlCall.xlcAlert, message, 2 /* xlAlertWarning */); }
            catch (Exception ex) { _log.Error("QueueStructureAlert failed", ex); }
        });
    }

    /// <summary>Point the workbook's current-sheet pointer at its live active
    /// sheet so <see cref="StateService.Get"/> projects the right sheet's profile.
    /// Cheap and idempotent — <see cref="StateService.SetCurrentSheet"/> only
    /// repaints when the active sheet actually changed.</summary>
    private void SetCurrentSheetToLive(Excel.Workbook wb)
        => _state.SetCurrentSheet(KeyOf(wb), ActiveSheetName(wb));

    /// <summary>The workbook's active sheet name, or null on a COM hiccup (the
    /// pointer then falls back to the workbook's default bucket). Read late-bound
    /// so it works for a worksheet or a chart sheet alike.</summary>
    private static string? ActiveSheetName(Excel.Workbook wb)
    {
        try
        {
            dynamic sheet = wb.ActiveSheet;
            return sheet is null ? null : (string)sheet.Name;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Best-effort delete of any external profile sidecar, called from
    /// <see cref="OnWorkbookAfterSave"/> once the in-workbook part is safely on
    /// disk — never before, so a failed save can't lose state.</summary>
    private void DeleteLegacySidecars(Excel.Workbook wb)
    {
        foreach (var path in LegacySidecarPaths(wb))
        {
            try { if (System.IO.File.Exists(path)) System.IO.File.Delete(path); }
            catch { /* best-effort cleanup */ }
        }
    }

    /// <summary>The external profile sidecar paths earlier builds wrote, newest
    /// first: <c>&lt;projectDir&gt;/.pyexcel/project.xml</c> then the loose
    /// <c>&lt;projectDir&gt;/pyexcel.project.xml</c>. Empty when the workbook has
    /// no local project folder.</summary>
    private System.Collections.Generic.IEnumerable<string> LegacySidecarPaths(Excel.Workbook wb)
    {
        var dir = ResolveProjectDir(wb);
        if (string.IsNullOrEmpty(dir)) yield break;
        yield return System.IO.Path.Combine(dir!, ".pyexcel", "project.xml");
        yield return System.IO.Path.Combine(dir!, "pyexcel.project.xml");
    }

    /// <summary>(Re)point the live script watcher at the active workbook's
    /// userScripts folder so the ribbon's Script dropdown tracks files appearing
    /// / disappearing without a re-activate. The watcher pushes an initial
    /// snapshot from its constructor, so this also populates the list now.
    ///
    /// <para>Elements load only when the project is <see cref="ProjectReadiness.Ready"/>
    /// (enabled AND structure validated): a not-yet-enabled or broken project has
    /// nothing trustworthy to surface, so we drop the watcher and publish an empty
    /// list — keeping validation and element-loading in lockstep with the button
    /// gate. The watcher (re)starts the moment the project becomes Ready, because
    /// Enable/Repair refreshes the health verdict and then fires
    /// <see cref="PyExcelServices.RequestScriptRefresh"/>.</para></summary>
    private void SyncScriptWatcher(Excel.Workbook wb)
    {
        string key = KeyOf(wb);
        var ready = PyExcelServices.Health.ReadinessOf(key, _state.GetProfile(key).Enabled)
                    == ProjectReadiness.Ready;
        var scriptsDir = ready ? ResolveScriptsDir(wb) : null;

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

    private void Guard(string handler, Action body)
    {
        try
        {
            body();
        }
        catch (Exception ex)
        {
            // Never let an exception cross back into Excel's event pump. Log to
            // the file so a handler fault is diagnosable from %TEMP%\PyExcel_Debug.log.
            _log.Error($"AppEventSink.{handler} failed", ex);
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
            _app.WorkbookAfterSave -= OnWorkbookAfterSave;
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
