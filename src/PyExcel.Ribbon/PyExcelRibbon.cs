using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Logging;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;
using PyExcel.Forms;
using PyExcel.State;

namespace PyExcel.Ribbon;

/// <summary>
/// PyExcel v2 ribbon — sole owner of the ribbon-callback surface declared in
/// <c>Resources/RibbonUI.xml</c>. Callback names match the v1 VBA in
/// <c>src/module/modRibbon.bas</c> verbatim so the XML can be diffed
/// line-for-line against v1.
/// </summary>
/// <remarks>
/// PHASE 1 SCOPE: this class exists to prove the ribbon loads. Only
/// <c>OnReadMe</c> is wired to real behaviour. Every other onAction is a
/// stub that logs and returns; every getter returns a safe default. Each
/// stub explicitly references the v1 line it must eventually match, so
/// porting in later phases is mechanical.
///
/// SAFE-1 (ribbon callbacks never touch the pipe) is structurally enforced
/// in this file: there is no field of type <c>KernelSupervisor</c>,
/// <c>NamedPipeClientStream</c>, or anything pipe-related. When the bridge
/// is added (Phase 2), the OnRunPython callback will enqueue work to a
/// background service via a thread-safe queue and return immediately —
/// it must not block, and it must not synchronously read from the pipe.
/// </remarks>
[ComVisible(true)]
public class PyExcelRibbon : ExcelRibbon
{
    private readonly ILog _log = new FileLog();
    private IRibbonUI? _ribbon;

    /// <summary>Loaded from the embedded RibbonUI.xml resource.</summary>
    public override string GetCustomUI(string RibbonID)
    {
        try
        {
            using var stream = typeof(PyExcelRibbon).Assembly
                .GetManifestResourceStream("PyExcel.Ribbon.Resources.RibbonUI.xml")
                ?? throw new InvalidOperationException(
                    "RibbonUI.xml embedded resource missing — build is malformed");
            using var reader = new StreamReader(stream);
            return reader.ReadToEnd();
        }
        catch (Exception ex)
        {
            _log.Error("GetCustomUI failed", ex);
            // Returning empty string causes Excel to render no ribbon tab —
            // strictly better than crashing the host process.
            return string.Empty;
        }
    }

    // -------------------------------------------------------------------------
    // Ribbon lifecycle
    // -------------------------------------------------------------------------

    public void RibbonOnLoad(IRibbonUI ribbon)
    {
        _ribbon = ribbon;
        _log.Info("Ribbon loaded");
        // Subscribe to state changes so the ribbon redraws when the
        // active workbook's state mutates. We queue Invalidate onto the
        // macro queue rather than calling it inline — state changes can
        // originate from a FileSystemWatcher worker thread, and
        // IRibbonUI is COM-affine.
        PyExcelServices.State.StateChanged += OnStateChanged;
        // Also repaint on error capture / clear so the "Show / Copy
        // Last Error" buttons enable as soon as a run fails (and disable
        // again when the slot is cleared). Same queue-then-invalidate
        // contract — the recorder runs on a background task.
        PyExcelServices.Errors.ErrorChanged += OnErrorChanged;
        // Let non-ribbon components ask for a repaint without reaching into
        // this class. The COM event sink uses it on WorkbookActivate, where
        // the active workbook key changes but nothing in the registry
        // mutates — so no StateChanged fires, yet every getter must
        // re-render against the newly-active workbook's state.
        PyExcelServices.RequestRibbonInvalidate = QueueInvalidate;
    }

    private void OnStateChanged(object? sender, StateChangedEventArgs e)
    {
        if (_ribbon is null) return;
        // Skip work if the change is to a workbook other than the active
        // one — most cells of the ribbon only render the active workbook.
        var activeKey = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (activeKey is not null
            && !string.Equals(activeKey, e.WorkbookKey, StringComparison.Ordinal))
        {
            return;
        }
        QueueInvalidate();
    }

    private void OnErrorChanged(object? sender, ErrorChangedEventArgs e)
    {
        if (_ribbon is null) return;
        // A null WorkbookKey is the global slot — surface those on every
        // workbook, since no specific workbook owns them.
        if (e.WorkbookKey is null)
        {
            QueueInvalidate();
            return;
        }
        var activeKey = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (activeKey is not null
            && !string.Equals(activeKey, e.WorkbookKey, StringComparison.Ordinal))
        {
            return;
        }
        QueueInvalidate();
    }

    /// <summary>Queue an <see cref="IRibbonUI.Invalidate"/> onto Excel's
    /// macro thread. Safe to call from any thread (FileSystemWatcher,
    /// COM event sink): <see cref="IRibbonUI"/> is COM-affine, so we never
    /// invalidate inline.</summary>
    private void QueueInvalidate()
    {
        if (_ribbon is null) return;
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() => _ribbon?.Invalidate());
        }
        catch (Exception ex)
        {
            _log.Error("Ribbon Invalidate queue failed", ex);
        }
    }

    /// <summary>Discard a user edit to a display-only ribbon control by
    /// re-reading the getters, so Excel replaces the typed text with the
    /// behind-the-scenes value. The Script / Input / Output fields are driven by
    /// the selected action, not typed in — only the Actions dropdown is
    /// interactive — so their onChange callbacks route here instead of storing.
    /// Uses the full <see cref="QueueInvalidate"/> (the path already proven to
    /// refresh these <c>getText</c> boxes) so the revert is reliable.</summary>
    private void RevertControl(IRibbonControl control) => QueueInvalidate();

    /// <summary>Read the state for the currently-active workbook from the
    /// in-memory registry, returning <see cref="WorkbookState.Empty"/> if no
    /// workbook is active so every getter has a well-defined value to read.
    ///
    /// <para>This does <b>no I/O</b>: restoring a workbook's saved profile from
    /// its embedded part is event-driven — the COM event sink populates the
    /// registry on open/activate (and for already-open workbooks at add-in load,
    /// via <c>RestoreOpenWorkbooks</c>). Keeping the render path pure-memory means
    /// a plain, non-PyExcel workbook costs nothing to draw (no per-repaint disk or
    /// COM probe), which is what keeps the ribbon cheap when PyExcel isn't enabled
    /// for the active workbook.</para></summary>
    private WorkbookState ActiveState()
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return WorkbookState.Empty("<no-workbook>");
        return PyExcelServices.State.Get(key);
    }

    public override object? LoadImage(string imageName)
    {
        // imageName="customLogo" — load the PNG shipped as an embedded
        // resource (LogicalName "customLogo.png" in PyExcel.Ribbon.csproj).
        // Returned as a System.Drawing.Bitmap, which Excel-DNA converts to
        // the IPictureDisp the ribbon expects.
        if (!string.Equals(imageName, "customLogo", StringComparison.Ordinal))
            return null;

        try
        {
            var assembly = typeof(PyExcelRibbon).Assembly;
            using var stream = assembly.GetManifestResourceStream("customLogo.png");
            if (stream is null) return null;
            // Clone off the stream so the Bitmap doesn't depend on it
            // staying open (the classic Bitmap(Stream) lifetime trap).
            using var fromStream = new System.Drawing.Bitmap(stream);
            return new System.Drawing.Bitmap(fromStream);
        }
        catch
        {
            // A missing/corrupt resource must never break ribbon load —
            // fall back to no image.
            return null;
        }
    }

    // -------------------------------------------------------------------------
    // getEnabled — every control's enabled-state derives from one readiness
    // verdict (ActiveReadiness), so the data controls and the Enable/Repair
    // button are exact complements and can never drift. Mirrors v1
    // RibbonIsEnabled in modRibbon.bas.
    // -------------------------------------------------------------------------

    /// <summary>The active workbook's consolidated readiness — the single gate every
    /// data control reads, so "is the project usable right now?" is decided in
    /// exactly one place. Combines the persisted <see cref="WorkbookState.Enabled"/>
    /// flag with the last on-disk structure check (recorded by the COM sink on
    /// open/activate, by Enable/Repair, and by the Run guard): a workbook is
    /// <see cref="ProjectReadiness.Ready"/> only when it is enabled AND its
    /// environment validated whole. Reads the active workbook key once so the render
    /// path's COM cost stays flat.</summary>
    private static ProjectReadiness ActiveReadiness()
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        var enabled = key is not null && PyExcelServices.State.Get(key).Enabled;
        return PyExcelServices.Health.ReadinessOf(key, enabled);
    }

    /// <summary>getEnabled shared by every data control (Run, Edit, the Script /
    /// Input / Output / Actions fields, Import / Export / Paste, Open Explorer, Read
    /// Me), mirroring v1 <c>RibbonIsEnabled</c>. Live only when the active workbook is
    /// <see cref="ProjectReadiness.Ready"/> — enabled for PyExcel AND its on-disk
    /// structure validated — so the user can't drive a half-installed project into a
    /// cryptic kernel error.</summary>
    public bool RibbonEnabled(IRibbonControl control)
        => ActiveReadiness() == ProjectReadiness.Ready;

    /// <summary>getEnabled for the "Enable" button — the logical complement of
    /// <see cref="RibbonEnabled"/>: clickable whenever the active workbook is not yet
    /// fully ready, i.e. either never enabled, or enabled with a missing/incomplete
    /// environment (in which case the same button doubles as "Repair" and
    /// re-provisions the missing venv/kernel). It greys out only once the workbook is
    /// enabled and its structure validated whole.</summary>
    public bool RibbonNotEnabled(IRibbonControl control)
        => ActiveReadiness() != ProjectReadiness.Ready;

    /// <summary>getEnabled for the "Update" button. PLACEHOLDER (Note 3):
    /// always false until the update mechanism and a launch-time update check
    /// land — see OnUpdate and ROADMAP.md (Phase 9).</summary>
    public bool RibbonUpdateAvailable(IRibbonControl control) => false;

    // -------------------------------------------------------------------------
    // Main group
    // -------------------------------------------------------------------------

    public void OnEnablePyExcel(IRibbonControl control)
    {
        // "Enable" turns a plain workbook into a PyExcel workbook: it runs the
        // full setup (create the project folders, provision the venv, extract
        // the kernel, install dependencies) and, on success, marks the workbook
        // enabled. Install and enable are deliberately one action (Note 3). The
        // ribbon greys this button once the workbook is enabled (getEnabled =
        // RibbonNotEnabled), so it can't be re-run by accident; flipping Enabled
        // fires StateChanged, which RibbonOnLoad turns into an
        // IRibbonUI.Invalidate so every getEnabled-gated control repaints.
        _log.Info("OnEnablePyExcel clicked");
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnEnablePyExcel: no active workbook"); return; }

            // Repair mode: an already-enabled workbook whose environment failed the
            // structure check (readiness == NeedsRepair, which already implies
            // enabled). Re-provision into its existing project folder — no
            // name/folder prompt, we already know where it lives.
            if (PyExcelServices.Health.ReadinessOf(key, PyExcelServices.State.Get(key).Enabled)
                == ProjectReadiness.NeedsRepair)
            {
                RepairActiveWorkbook(key);
                return;
            }

            // A brand-new workbook has never been saved (Workbook.Path is empty),
            // so it has no folder to anchor the project to. Rather than refuse,
            // ask for a name and save it into the folder the user picks next.
            var workbookDir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
            string? newWorkbookName = null;
            if (string.IsNullOrEmpty(workbookDir))
            {
                newWorkbookName = PromptForWorkbookName();
                if (newWorkbookName is null) { _log.Info("OnEnablePyExcel: name prompt cancelled"); return; }
            }

            // Open the folder browser at the workbook's own folder so the user
            // starts where the workbook lives; fall back to Documents for a
            // new/unsaved (or cloud-URL) workbook that has no local folder.
            var browseStart = !string.IsNullOrEmpty(workbookDir) && Directory.Exists(workbookDir!)
                ? workbookDir!
                : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var projectDir = PickProjectDirectory(browseStart);
            if (projectDir is null) { _log.Info("OnEnablePyExcel: directory pick cancelled"); return; }

            // For a new workbook, save it into the chosen folder first so its
            // persisted state has a file to live in, then re-key off the saved
            // path (Save As promotes the workbook to a path-based key).
            if (newWorkbookName is not null)
            {
                var savePath = Path.Combine(projectDir, EnsureXlsxExtension(newWorkbookName));
                if (!SaveActiveWorkbookAs(savePath))
                {
                    LogDisplay.WriteLine($"Enable: couldn't save the workbook to '{savePath}'.");
                    return;
                }
                key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
                if (key is null) { _log.Info("OnEnablePyExcel: workbook key lost after save"); return; }
            }

            // Remember the choice so Setup and the runtime kernel both use it.
            PyExcelServices.State.SetProjectDir(key, projectDir);

            var success = SetupForm.Run(ExcelWindowOwner(), projectDir, _log);
            if (success == true)
            {
                PyExcelServices.State.SetEnabled(key, true);
                // The workbook is the single store, and its embedded profile part
                // is written by the save hook. PyExcel state lives outside the
                // cell grid, so enabling doesn't dirty the workbook on its own —
                // mark it dirty and save now so the enabled state is flushed into
                // the part and survives a close-without-save. (Enable is a heavy,
                // one-time, explicitly-clicked action, so saving as part of it is
                // expected; lighter edits below just mark dirty.)
                MarkWorkbookDirty();
                SaveActiveWorkbook();
                // Record the freshly-provisioned structure as healthy so the Enable
                // button greys out (it doubles as Repair when the structure breaks).
                PyExcelServices.Health.Set(key, PyExcel.State.ProjectStructureValidator.Validate(projectDir));
                // Surface the scaffolded example.py now, and start the live
                // watcher on the just-created userScripts folder (Enable fires
                // no WorkbookActivate, so the sink wouldn't otherwise start it).
                RefreshAvailableScripts(key);
                PyExcelServices.RequestScriptRefresh?.Invoke();
                _log.Info($"OnEnablePyExcel: workbook '{key}' set up at '{projectDir}' and enabled");
            }
            else
            {
                _log.Info($"OnEnablePyExcel: setup did not complete; '{key}' left disabled");
            }
        }
        catch (Exception ex)
        {
            _log.Error("OnEnablePyExcel failed", ex);
            LogDisplay.WriteLine($"Enable: {ex.Message}");
        }
    }

    /// <summary>Re-provision an already-enabled workbook whose open-time structure
    /// check failed: run Setup against the known project folder (no name/folder
    /// prompt), then re-validate and repaint so the Enable/Repair button greys out
    /// once the environment is whole again.</summary>
    private void RepairActiveWorkbook(string key)
    {
        try
        {
            var projectDir = ResolveProjectDir();
            if (string.IsNullOrEmpty(projectDir))
            {
                LogDisplay.WriteLine("Repair: no project folder is known for this workbook.");
                return;
            }
            _log.Info($"OnEnablePyExcel: repairing '{key}' at '{projectDir}'");
            var success = SetupForm.Run(ExcelWindowOwner(), projectDir!, _log);
            if (success == true)
            {
                PyExcelServices.Health.Set(key, PyExcel.State.ProjectStructureValidator.Validate(projectDir));
                RefreshAvailableScripts(key);
                PyExcelServices.RequestScriptRefresh?.Invoke();
                QueueInvalidate();
                _log.Info($"OnEnablePyExcel: repair complete for '{key}'");
            }
            else
            {
                _log.Info($"OnEnablePyExcel: repair did not complete for '{key}'");
            }
        }
        catch (Exception ex)
        {
            _log.Error("RepairActiveWorkbook failed", ex);
            LogDisplay.WriteLine($"Repair: {ex.Message}");
        }
    }

    public void OnUpdate(IRibbonControl control)
    {
        // PLACEHOLDER (Note 3). The update path — refresh the extracted kernel
        // and re-sync dependencies, plus a launch-time "is a newer build
        // available?" check that would drive RibbonUpdateAvailable — isn't built
        // yet. The button is greyed (RibbonUpdateAvailable returns false), so
        // this is normally unreachable; it logs if invoked. Tracked as an open
        // item in ROADMAP.md (Phase 9).
        _log.Info("OnUpdate clicked (placeholder — update not yet implemented)");
        LogDisplay.WriteLine("[PyExcel] Update isn't available yet.");
    }

    public void OnOpenExplorer(IRibbonControl control)
    {
        _log.Info("OnOpenExplorer clicked");
        try
        {
            // Open the project directory (where the venv, kernel, and
            // userScripts live) — the dedicated folder chosen on Enable, else
            // the workbook's own folder.
            var dir = ResolveProjectDir();
            if (string.IsNullOrEmpty(dir))
            {
                LogDisplay.WriteLine(
                    "Open Explorer: the active workbook hasn't been saved yet — " +
                    "no directory to open.");
                return;
            }
            ShellLauncher.OpenInExplorer(dir!);
        }
        catch (Exception ex)
        {
            _log.Error("OnOpenExplorer failed", ex);
            LogDisplay.WriteLine($"Open Explorer: {ex.Message}");
        }
    }

    public void OnReadMe(IRibbonControl control)
    {
        _log.Info("OnReadMe clicked");
        try
        {
            // Look for a README at the project folder (where Setup scaffolds it)
            // first, then next to the workbook. Setup writes one on Enable, but
            // workbooks enabled before that existed won't have it — so if we have
            // a project folder and there's no README, write a default one now so
            // the button always opens something useful instead of doing nothing.
            var projectDir = ResolveProjectDir();
            var workbookDir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;

            var readme = FirstExistingReadme(projectDir, workbookDir);
            if (readme is null && !string.IsNullOrEmpty(projectDir))
            {
                var candidate = Path.Combine(projectDir!, "README.md");
                try
                {
                    Directory.CreateDirectory(projectDir!);
                    File.WriteAllText(candidate, DefaultReadme);
                    readme = candidate;
                }
                catch (Exception ex)
                {
                    _log.Error("OnReadMe: couldn't create README", ex);
                }
            }

            if (readme is not null && File.Exists(readme))
            {
                ShellLauncher.Open(readme);
                return;
            }

            const string text =
                "PyExcel\n\n" +
                "No project folder is set up for this workbook yet. " +
                "Click Enable first — that creates the project folder " +
                "(with a README and an example script) for this workbook.";
            XlCall.Excel(XlCall.xlcAlert, text, 2 /* xlAlertWarning */);
        }
        catch (Exception ex)
        {
            _log.Error("OnReadMe failed", ex);
        }
    }

    /// <summary>The first <c>README.md</c> that exists at the project folder or
    /// the workbook folder, or null if neither has one.</summary>
    private static string? FirstExistingReadme(string? projectDir, string? workbookDir)
    {
        foreach (var dir in new[] { projectDir, workbookDir })
        {
            if (string.IsNullOrEmpty(dir)) continue;
            var path = Path.Combine(dir!, "README.md");
            if (File.Exists(path)) return path;
        }
        return null;
    }

    /// <summary>Minimal README written on demand when a project folder has none
    /// (e.g. a workbook enabled before Setup started scaffolding one).</summary>
    private const string DefaultReadme =
        "# PyExcel project\n\n" +
        "This folder holds the PyExcel environment for your workbook " +
        "(`userScripts/`, `.pyexcel-venv/`, `.pyexcel-kernel/`).\n\n" +
        "- Pick a script in the ribbon's Script box and click Edit to change it.\n" +
        "- Click Add to bind a script to input/output ranges, then Run.\n" +
        "- `print()` output and errors show in the log window.\n\n" +
        "Your actions and settings are saved inside the workbook.\n";

    // -------------------------------------------------------------------------
    // Python group
    // -------------------------------------------------------------------------

    public void OnRunPython(IRibbonControl control)
    {
        _log.Info("OnRunPython clicked");
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnRunPython: no active workbook"); return; }

            // Defensive readiness gate. getEnabled normally greys Run when the
            // environment is incomplete, but Excel can act on a stale ribbon cache,
            // and the venv/kernel can vanish while the workbook stays open. Re-check
            // the on-disk structure now (cheap, file-only) and refuse with a clear,
            // actionable message — repainting so Run greys and Enable/Repair lights —
            // rather than booting a kernel that would fail with a cryptic error.
            if (!EnsureStructureReady(key, "Run")) return;

            // RangeRunner reads the input ranges synchronously on this
            // (main) thread, then dispatches the kernel exchange to a
            // background task and writes the result back via QueueAsMacro —
            // so this callback returns promptly and never blocks on the
            // pipe (SAFE-1).
            var state = PyExcelServices.State.Get(key);
            // Supply a modeless progress dialog with a working Cancel; the
            // factory is invoked on this (main) thread so the form pumps
            // while the kernel runs on the background task.
            PyExcel.Excel.RangeRunner.RunActiveScript(
                state,
                progressFactory: () =>
                    ProgressForm.StartModeless(ExcelWindowOwner(), "Running Python…"),
                // On failure, pop the error in a modal dialog that stays up until
                // the user dismisses it — a Python traceback shouldn't just scroll
                // by in the log window where it's easy to miss.
                errorDisplay: message =>
                    ErrorDisplayForm.Open(ExcelWindowOwner(), "PyExcel — Run failed", message));
        }
        catch (Exception ex)
        {
            _log.Error("OnRunPython failed", ex);
        }
    }

    /// <summary>Re-validate the active workbook's on-disk project structure right
    /// before a kernel-bound action, and record the verdict in
    /// <see cref="PyExcelServices.Health"/> so the ribbon's readiness gate stays
    /// current. Returns <see langword="true"/> when the structure is whole.
    /// Otherwise it repaints the ribbon (Run greys, Enable/Repair lights), surfaces an
    /// actionable modal naming what's missing, and returns <see langword="false"/> —
    /// the caller must not proceed. This is the last-line defence behind the
    /// getEnabled gate, for the cases Excel's ribbon cache can't catch (a stale paint,
    /// or the environment deleted mid-session).</summary>
    private bool EnsureStructureReady(string key, string action)
    {
        var projectDir = ResolveProjectDir();
        var check = PyExcel.State.ProjectStructureValidator.Validate(projectDir);
        PyExcelServices.Health.Set(key, check);
        if (check.Ok) return true;

        QueueInvalidate();
        var body = new System.Text.StringBuilder();
        body.Append(action).Append(" can't start — this workbook's PyExcel ");
        body.Append("environment is missing or incomplete:\n\n");
        foreach (var m in check.Missing) body.Append("  • ").Append(m).Append('\n');
        body.Append("\nOpen the PyExcel tab and click Enable to reinstall it.");
        ErrorDisplayForm.Open(
            ExcelWindowOwner(), "PyExcel — environment incomplete", body.ToString());
        _log.Info($"{action}: blocked — structure incomplete at '{projectDir}': " +
                  string.Join(", ", check.Missing));
        return false;
    }

    public void OnEditPython(IRibbonControl control)
    {
        _log.Info("OnEditPython clicked");
        try
        {
            var script = ActiveState().SelectedScript;
            if (string.IsNullOrEmpty(script))
            {
                LogDisplay.WriteLine(
                    "Edit Python: no script is selected. " +
                    "Pick one from the ribbon's Script dropdown first.");
                return;
            }
            var dir = ResolveProjectDir();
            if (string.IsNullOrEmpty(dir))
            {
                LogDisplay.WriteLine(
                    "Edit Python: the active workbook hasn't been saved yet — " +
                    "save the workbook first so the userScripts/ folder is " +
                    "located on disk.");
                return;
            }
            // Convention: scripts live under <projectDir>/userScripts/<name> —
            // the dedicated folder chosen on Enable, else the workbook folder.
            // Setup and the "New…" button scaffold into the same root.
            var path = Path.Combine(dir!, "userScripts", script!);
            if (!File.Exists(path))
            {
                LogDisplay.WriteLine($"Edit Python: file not found at '{path}'.");
                return;
            }
            ShellLauncher.Open(path);
        }
        catch (Exception ex)
        {
            _log.Error("OnEditPython failed", ex);
            LogDisplay.WriteLine($"Edit Python: {ex.Message}");
        }
    }

    public int GetScriptCount(IRibbonControl control) => ActiveState().AvailableScripts.Count;

    public string GetScriptLabel(IRibbonControl control, int index)
    {
        var scripts = ActiveState().AvailableScripts;
        return index >= 0 && index < scripts.Count ? scripts[index] : string.Empty;
    }

    public string GetScriptText(IRibbonControl control) => ActiveState().SelectedScript ?? string.Empty;

    // Display-only: the Script box mirrors the selected action; reject hand-edits.
    public void OnScriptChange(IRibbonControl control, string text) => RevertControl(control);

    public string GetPyInput(IRibbonControl control) => ActiveState().PyInput ?? string.Empty;

    // Display-only: the Input box mirrors the selected action; reject hand-edits.
    public void OnPyInputChange(IRibbonControl control, string text) => RevertControl(control);

    public string GetPyOutput(IRibbonControl control) => ActiveState().PyOutput ?? string.Empty;

    // Display-only: the Output box mirrors the selected action; reject hand-edits.
    public void OnPyOutputChange(IRibbonControl control, string text) => RevertControl(control);

    public int GetActionCount(IRibbonControl control) => ActiveState().Actions.Count;

    public string GetActionLabel(IRibbonControl control, int index)
    {
        var actions = ActiveState().Actions;
        return index >= 0 && index < actions.Count ? actions[index].Name : string.Empty;
    }

    public string GetActionText(IRibbonControl control)
        => ActiveState().SelectedAction?.Name ?? string.Empty;

    public void OnActionChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        if (string.IsNullOrEmpty(text))
        {
            PyExcelServices.State.SetSelectedAction(key, null);
            return;
        }
        // Selecting a saved action loads it into the Script / Input / Output
        // boxes (which the Run button reads), so picking an action makes it
        // runnable — not just a name in the combo.
        foreach (var a in PyExcelServices.State.Get(key).Actions)
        {
            if (string.Equals(a.Name, text, StringComparison.Ordinal))
            {
                PyExcelServices.State.LoadAction(key, a);
                return;
            }
        }
        PyExcelServices.State.SetSelectedAction(key, text);
    }

    public void OnAddAction(IRibbonControl control)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) { _log.Info("OnAddAction: no active workbook"); return; }

        // Re-scan userScripts so the form's Script list reflects what's on disk.
        RefreshAvailableScripts(key);
        var state = PyExcelServices.State.Get(key);
        var result = EditActionForm.Prompt(
            ExcelWindowOwner(),
            state.AvailableScripts,
            ActionNames(state),
            existing: null,
            rangePicker: PickRangeNative,
            userScriptsDirectory: UserScriptsDir());
        if (result is null) { _log.Info("OnAddAction: cancelled"); return; }

        PyExcelServices.State.AddAction(key, result);
        // Load it into the run boxes so the just-saved action is ready to Run.
        PyExcelServices.State.LoadAction(key, result);
        MarkWorkbookDirty();
        _log.Info($"OnAddAction: saved '{result.Name}' to workbook '{key}'");
    }

    public void OnEditAction(IRibbonControl control)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) { _log.Info("OnEditAction: no active workbook"); return; }

        // Re-scan userScripts so the form's Script list reflects what's on disk.
        RefreshAvailableScripts(key);
        var state = PyExcelServices.State.Get(key);
        var existing = state.SelectedAction;
        if (existing is null) { _log.Info("OnEditAction: no action selected"); return; }

        var result = EditActionForm.Prompt(
            ExcelWindowOwner(),
            state.AvailableScripts,
            ActionNames(state),
            existing,
            rangePicker: PickRangeNative,
            userScriptsDirectory: UserScriptsDir());
        if (result is null) { _log.Info("OnEditAction: cancelled"); return; }

        // AddAction upserts by name, so a rename would leave the original
        // entry behind under its old name — drop it first when the name
        // changed. (The validator already guaranteed the new name doesn't
        // collide with a *different* action.)
        if (!string.Equals(existing.Name, result.Name, StringComparison.Ordinal))
            PyExcelServices.State.DeleteAction(key, existing.Name);
        PyExcelServices.State.AddAction(key, result);
        // Reflect the edited action in the run boxes (Script / Input / Output).
        PyExcelServices.State.LoadAction(key, result);
        MarkWorkbookDirty();
        _log.Info($"OnEditAction: saved '{result.Name}' to workbook '{key}'");
    }

    /// <summary>The existing action names, in order, for the EditAction
    /// dialog's duplicate-name check.</summary>
    private static System.Collections.Generic.IReadOnlyList<string> ActionNames(WorkbookState state)
    {
        var names = new System.Collections.Generic.List<string>(state.Actions.Count);
        foreach (var a in state.Actions) names.Add(a.Name);
        return names;
    }

    /// <summary>Wraps Excel's main window so a modal dialog is owned by it
    /// — never lost behind Excel or off-screen (the v1 hide hack is gone).</summary>
    private static System.Windows.Forms.IWin32Window ExcelWindowOwner()
        => new ExcelWindow(ExcelDnaUtil.WindowHandle);

    private sealed class ExcelWindow : System.Windows.Forms.IWin32Window
    {
        public ExcelWindow(IntPtr handle) => Handle = handle;
        public IntPtr Handle { get; }
    }

    /// <summary>The active workbook's <c>userScripts</c> folder under its
    /// resolved project directory, or null if the workbook hasn't been saved.
    /// Passed to the EditAction dialog so its "New…" button scaffolds a script
    /// where Setup and the runtime look (Note 2).</summary>
    private static string? UserScriptsDir()
    {
        var projectDir = ResolveProjectDir();
        return string.IsNullOrEmpty(projectDir) ? null : Path.Combine(projectDir!, "userScripts");
    }

    /// <summary>Re-scan the active workbook's userScripts folder and push the
    /// result into <see cref="WorkbookState.AvailableScripts"/>, so the ribbon's
    /// Script dropdown and the Add/Edit form's script list both reflect what's
    /// actually on disk. Nothing else populates that list in-process, so this is
    /// called on Enable and whenever the action form is opened.</summary>
    private static void RefreshAvailableScripts(string key)
        => PyExcelServices.State.SetAvailableScripts(
            key, ScriptDirectoryWatcher.Snapshot(UserScriptsDir()));

    /// <summary>The active workbook's effective project directory: the dedicated
    /// folder the user chose on Enable (saved in state) if set, else the
    /// workbook-derived default from <see cref="PyExcel.Common.ProjectDirectory"/>.
    /// Null when no workbook is active or it's unsaved. Mirrors the rule
    /// KernelHost uses so the ribbon and runtime agree on the directory.</summary>
    private static string? ResolveProjectDir()
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        var workbookDir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
        var stored = key is null ? null : PyExcelServices.State.Get(key).ProjectDir;
        return string.IsNullOrEmpty(stored)
            ? PyExcel.Common.ProjectDirectory.Resolve(workbookDir)
            : stored;
    }

    /// <summary>Prompt for a dedicated project folder via the native folder
    /// browser, defaulting to <paramref name="defaultDir"/>. Returns the chosen
    /// absolute path, or null if the user cancelled. Raises Excel first so the
    /// browser isn't lost behind the Excel window.</summary>
    private string? PickProjectDirectory(string? defaultDir)
    {
        try
        {
            try { SetForegroundWindow(ExcelDnaUtil.WindowHandle); } catch { /* best-effort */ }
            using var dlg = new System.Windows.Forms.FolderBrowserDialog
            {
                Description =
                    "Choose a dedicated folder for this workbook's PyExcel project. " +
                    "Its Python environment (.pyexcel-venv), kernel, and userScripts " +
                    "folder are created here.",
                ShowNewFolderButton = true,
            };
            if (!string.IsNullOrEmpty(defaultDir) && Directory.Exists(defaultDir))
                dlg.SelectedPath = defaultDir!;

            var result = dlg.ShowDialog(ExcelWindowOwner());
            return result == System.Windows.Forms.DialogResult.OK &&
                   !string.IsNullOrWhiteSpace(dlg.SelectedPath)
                ? dlg.SelectedPath
                : null;
        }
        catch (Exception ex)
        {
            _log.Error("PickProjectDirectory failed", ex);
            return null;
        }
    }

    /// <summary>Prompt (via Excel's own InputBox) for a name for a brand-new,
    /// never-saved workbook. Returns the sanitised name (no extension), or null
    /// if the user cancelled or left it blank. Defaults to the workbook's current
    /// caption (e.g. "Book1").</summary>
    private string? PromptForWorkbookName()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            string current = string.Empty;
            try { current = (string)app.ActiveWorkbook.Name; } catch { /* best-effort default */ }

            try { SetForegroundWindow(ExcelDnaUtil.WindowHandle); } catch { /* best-effort */ }

            // InputBox Type 2 = text. Cancel returns the Boolean False; OK returns
            // the typed string (empty if the user cleared it).
            object result = app.InputBox(
                "This workbook isn't saved yet. Enter a name — PyExcel will save " +
                "it into the folder you choose next.",
                "PyExcel — name this workbook",
                current,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);

            if (result is bool) return null; // cancelled
            var name = (result as string)?.Trim();
            if (string.IsNullOrEmpty(name)) return null;

            // Strip anything that can't live in a filename so the later SaveAs
            // can't throw on the path.
            foreach (var c in Path.GetInvalidFileNameChars())
                name = name!.Replace(c.ToString(), string.Empty);
            name = name!.Trim();
            return string.IsNullOrEmpty(name) ? null : name;
        }
        catch (Exception ex)
        {
            _log.Error("PromptForWorkbookName failed", ex);
            return null;
        }
    }

    /// <summary>Append a <c>.xlsx</c> extension to <paramref name="name"/> unless
    /// it already has one (case-insensitive).</summary>
    private static string EnsureXlsxExtension(string name)
        => name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ? name : name + ".xlsx";

    /// <summary>Save the active (new, unsaved) workbook to <paramref name="path"/>
    /// as a macro-free <c>.xlsx</c> via COM. Returns false (logged) on failure.
    /// SaveAs updates the workbook's FullName, so the WorkbookContext reports the
    /// new path — and a path-based key — immediately afterwards.</summary>
    private bool SaveActiveWorkbookAs(string path)
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            // FileFormat 51 = xlOpenXMLWorkbook (.xlsx).
            app.ActiveWorkbook.SaveAs(path, 51);
            _log.Info($"SaveActiveWorkbookAs: saved new workbook to '{path}'");
            return true;
        }
        catch (Exception ex)
        {
            _log.Error($"SaveActiveWorkbookAs failed for '{path}'", ex);
            return false;
        }
    }

    /// <summary>Mark the active workbook as needing a save, so a subsequent
    /// Ctrl+S (or the close prompt) actually re-writes the file and flushes
    /// PyExcel's state into its CustomXMLPart. PyExcel state (enabled flag,
    /// actions, field bindings) lives outside the cell grid, so mutating it
    /// doesn't dirty the workbook on its own — without this, Excel can treat
    /// the save as a no-op and the user's enable/actions are never persisted.
    /// Best-effort and late-bound; a failure must never break the mutation that
    /// just happened.</summary>
    private void MarkWorkbookDirty()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb is not null) wb.Saved = false;
        }
        catch (Exception ex)
        {
            _log.Error("MarkWorkbookDirty failed", ex);
        }
    }

    /// <summary>Save the active workbook now via COM, so a state change that
    /// lives outside the cell grid (notably Enable) is flushed to the workbook's
    /// embedded profile part and survives a close-without-save. The save fires the
    /// event sink's <c>WorkbookBeforeSave</c>, which is what actually writes the
    /// part. Late-bound and best-effort: on a save failure the change stays in
    /// the in-memory registry (and the workbook stays marked dirty), so the
    /// user's next manual save still persists it.</summary>
    private void SaveActiveWorkbook()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic wb = app.ActiveWorkbook;
            if (wb is not null) wb.Save();
        }
        catch (Exception ex)
        {
            _log.Error("SaveActiveWorkbook failed", ex);
        }
    }

    /// <summary>Bring a window to the foreground. Used to raise Excel's main
    /// window before the native range picker so the picker isn't drawn behind
    /// Excel. Best-effort: the OS can refuse the foreground change.</summary>
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    /// <summary>
    /// Excel's NATIVE range picker (Application.InputBox with Type:=8): shows
    /// the collapsible "select a range" box so the user can click/drag on the
    /// sheet, and returns the chosen range's address as <c>Sheet!A1:B2</c>.
    /// Returns null if the user cancels or the call fails. Injected into the
    /// dialogs as the range-pick delegate so PyExcel.Forms stays COM-free.
    /// </summary>
    private string? PickRangeNative(string? initial)
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            // Pre-fill the picker with the field's current value, or — when it's
            // blank — the user's current sheet selection, so the common "I've
            // already selected it" case needs no extra drag.
            var seed = string.IsNullOrEmpty(initial) ? CurrentSelectionAddress() : initial;

            // Raise Excel to the foreground so the native picker shows on top of
            // the Excel window, not behind it — the dialog that launched the pick
            // is already hidden by RangePick.OnSheet. Best-effort (the OS can
            // refuse the foreground change), so failures are swallowed.
            try { SetForegroundWindow(ExcelDnaUtil.WindowHandle); } catch { /* best-effort */ }

            // Application.InputBox(Prompt, Title, Default, Left, Top, HelpFile,
            // HelpContextID, Type). Type 8 = a cell/range reference: Excel shows
            // its collapsible range selector. Cancel returns the Boolean False;
            // a pick returns a Range object.
            object box = app.InputBox(
                "Select a range, then click OK.",
                "PyExcel — pick a range",
                seed ?? string.Empty,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);

            if (box is bool) return null; // user cancelled

            dynamic range = box;
            string address = (string)range.Address[false, false];
            string sheet = (string)range.Worksheet.Name;
            return $"{sheet}!{address}";
        }
        catch (Exception ex)
        {
            _log.Error("PickRangeNative failed", ex);
            return null;
        }
    }

    /// <summary>The active selection's address as <c>Sheet!A1:B2</c>, used to
    /// pre-seed the native picker. Returns null when the selection isn't a range
    /// (a chart, a shape) or COM is unhappy — the picker then just opens with the
    /// field's existing text (or empty).</summary>
    private static string? CurrentSelectionAddress()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            dynamic selection = app.Selection;
            string address = (string)selection.Address[false, false];
            string sheet = (string)selection.Worksheet.Name;
            return $"{sheet}!{address}";
        }
        catch
        {
            return null;
        }
    }

    public void OnDeleteAction(IRibbonControl control)
    {
        // Unlike Add/Edit, Delete needs no form — we can wire it now.
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) { _log.Info("OnDeleteAction: no active workbook"); return; }
        var selected = PyExcelServices.State.Get(key).SelectedActionName;
        if (selected is null) { _log.Info("OnDeleteAction: no action selected"); return; }
        PyExcelServices.State.DeleteAction(key, selected);
        MarkWorkbookDirty();
        _log.Info($"OnDeleteAction: removed '{selected}' from workbook '{key}'");
    }

    // -------------------------------------------------------------------------
    // Import group — text fields persisted through StateService; the
    // ImportService button handler is the Phase 5 COM-bound follow-up
    // tracked in modRibbon.bas:593.
    // -------------------------------------------------------------------------

    public void OnImport(IRibbonControl control)
    {
        _log.Info("OnImport clicked");
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnImport: no active workbook"); return; }
            // The chooser runs on the macro thread, where the import opens
            // the workbook — show the picker owned by Excel's window.
            PyExcel.Excel.ImportService.RunActiveImport(
                PyExcelServices.State.Get(key),
                sheets => SheetPickerForm.Prompt(ExcelWindowOwner(), sheets, preselected: null));
        }
        catch (Exception ex)
        {
            _log.Error("OnImport failed", ex);
        }
    }

    public string GetImportInput(IRibbonControl control) => ActiveState().ImportInput ?? string.Empty;
    public void OnImportInputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetImportInput(key, text);
    }

    public string GetImportOutput(IRibbonControl control) => ActiveState().ImportOutput ?? string.Empty;
    public void OnImportOutputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetImportOutput(key, text);
    }

    public void OnEditImport(IRibbonControl control)
    {
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnEditImport: no active workbook"); return; }
            var state = PyExcelServices.State.Get(key);
            var result = EditIoForm.PromptImport(
                ExcelWindowOwner(),
                state.ImportInput,
                state.ImportOutput,
                PyExcelServices.WorkbookContext.CurrentWorkbookDirectory,
                PickRangeNative);
            if (result is null) { _log.Info("OnEditImport: cancelled"); return; }
            PyExcelServices.State.SetImportInput(key, result.Input);
            PyExcelServices.State.SetImportOutput(key, result.Output);
            _log.Info($"OnEditImport: saved for workbook '{key}'");
        }
        catch (Exception ex)
        {
            _log.Error("OnEditImport failed", ex);
        }
    }

    // -------------------------------------------------------------------------
    // Export group
    // -------------------------------------------------------------------------

    public void OnExport(IRibbonControl control)
    {
        _log.Info("OnExport clicked");
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnExport: no active workbook"); return; }
            PyExcel.Excel.ExportService.RunActiveExport(PyExcelServices.State.Get(key));
        }
        catch (Exception ex)
        {
            _log.Error("OnExport failed", ex);
        }
    }

    public string GetExportInput(IRibbonControl control) => ActiveState().ExportInput ?? string.Empty;
    public void OnExportInputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetExportInput(key, text);
    }

    public string GetExportOutput(IRibbonControl control) => ActiveState().ExportOutput ?? string.Empty;
    public void OnExportOutputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetExportOutput(key, text);
    }

    public void OnEditExport(IRibbonControl control)
    {
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnEditExport: no active workbook"); return; }
            var state = PyExcelServices.State.Get(key);
            var result = EditIoForm.PromptExport(
                ExcelWindowOwner(),
                state.ExportInput,
                state.ExportOutput,
                PyExcelServices.WorkbookContext.CurrentWorkbookDirectory,
                PickRangeNative);
            if (result is null) { _log.Info("OnEditExport: cancelled"); return; }
            PyExcelServices.State.SetExportInput(key, result.Input);
            PyExcelServices.State.SetExportOutput(key, result.Output);
            _log.Info($"OnEditExport: saved for workbook '{key}'");
        }
        catch (Exception ex)
        {
            _log.Error("OnEditExport failed", ex);
        }
    }

    public void OnExportWizard(IRibbonControl control)
    {
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnExportWizard: no active workbook"); return; }
            var state = PyExcelServices.State.Get(key);

            // Seed the wizard's first row from the single-export fields if set.
            System.Collections.Generic.IReadOnlyList<PyExcel.Excel.ExportJob>? seed = null;
            if (!string.IsNullOrWhiteSpace(state.ExportInput) ||
                !string.IsNullOrWhiteSpace(state.ExportOutput))
            {
                seed = new[]
                {
                    new PyExcel.Excel.ExportJob(state.ExportInput ?? string.Empty,
                                                state.ExportOutput ?? string.Empty),
                };
            }

            var jobs = ExportWizardForm.Prompt(
                ExcelWindowOwner(), seed,
                PyExcelServices.WorkbookContext.CurrentWorkbookDirectory);
            if (jobs is null) { _log.Info("OnExportWizard: cancelled"); return; }

            PyExcel.Excel.ExportService.RunBatch(
                jobs, PyExcelServices.WorkbookContext.CurrentWorkbookDirectory);
            _log.Info($"OnExportWizard: running {jobs.Count} export(s)");
        }
        catch (Exception ex)
        {
            _log.Error("OnExportWizard failed", ex);
        }
    }

    // -------------------------------------------------------------------------
    // Paste group
    // -------------------------------------------------------------------------

    public void OnPaste(IRibbonControl control)
    {
        _log.Info("OnPaste clicked");
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnPaste: no active workbook"); return; }
            PyExcel.Excel.PasteService.RunActivePaste(
                PyExcelServices.State.Get(key),
                orientationChooser: () => OrientationForm.Prompt(ExcelWindowOwner()));
        }
        catch (Exception ex)
        {
            _log.Error("OnPaste failed", ex);
        }
    }

    public string GetPasteOutput(IRibbonControl control) => ActiveState().PasteOutput ?? string.Empty;
    public void OnPasteOutputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetPasteOutput(key, text);
    }

    public void OnEditPaste(IRibbonControl control)
    {
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnEditPaste: no active workbook"); return; }
            var state = PyExcelServices.State.Get(key);
            var result = EditIoForm.PromptPaste(
                ExcelWindowOwner(), state.PasteOutput, PickRangeNative);
            if (result is null) { _log.Info("OnEditPaste: cancelled"); return; }
            PyExcelServices.State.SetPasteOutput(key, result.Output);
            _log.Info($"OnEditPaste: saved for workbook '{key}'");
        }
        catch (Exception ex)
        {
            _log.Error("OnEditPaste failed", ex);
        }
    }

    // -------------------------------------------------------------------------
    // Errors group — Show / Copy Last Error. Surfaces what the kernel
    // returned (or a host fault) without forcing the user to dig through
    // Excel-DNA's LogDisplay window.
    // -------------------------------------------------------------------------

    /// <summary>getEnabled for the error-group buttons: true iff
    /// <see cref="ErrorService"/> has something to show for the active
    /// workbook (or in the global slot).</summary>
    public bool RibbonHasError(IRibbonControl control)
    {
        try
        {
            return PyExcelServices.Errors.GetLast(ActiveKey()) is not null;
        }
        catch
        {
            // A faulty getEnabled would gray the button silently; better
            // to leave it enabled and surface the issue on click.
            return true;
        }
    }

    /// <summary>Show the last error in a modal, Excel-owned dialog that is
    /// always brought to the front. We deliberately do <em>not</em> route
    /// this through Excel-DNA's <see cref="LogDisplay"/>:
    /// <c>LogDisplay.Show()</c> is a no-op once the log window is already
    /// open behind Excel (and never raises it), so the button used to appear
    /// to do nothing. The formatted block is still written to LogDisplay for
    /// history, but the owned dialog is what guarantees the user sees it.</summary>
    public void OnShowLastError(IRibbonControl control)
    {
        try
        {
            var record = PyExcelServices.Errors.GetLast(ActiveKey());
            var body = record is null
                ? "No errors recorded."
                : record.FormatForClipboard();
            ErrorDisplayForm.Open(ExcelWindowOwner(), "PyExcel — Last Error", body);
        }
        catch (Exception ex)
        {
            _log.Error("OnShowLastError failed", ex);
        }
    }

    /// <summary>Copy the last error's formatted block to the clipboard
    /// so the user can paste it into a bug report. No-op (and a single
    /// LogDisplay note) when no error is on file.</summary>
    public void OnCopyLastError(IRibbonControl control)
    {
        try
        {
            var record = PyExcelServices.Errors.GetLast(ActiveKey());
            if (record is null)
            {
                LogDisplay.WriteLine("[PyExcel] No errors recorded — nothing to copy.");
                return;
            }
            // System.Windows.Forms.Clipboard requires an STA thread.
            // Ribbon callbacks run on Excel's main thread, which is STA.
            System.Windows.Forms.Clipboard.SetText(record.FormatForClipboard());
            LogDisplay.WriteLine("[PyExcel] Last error copied to clipboard.");
        }
        catch (Exception ex)
        {
            _log.Error("OnCopyLastError failed", ex);
        }
    }

    /// <summary>Active workbook key, or <see langword="null"/> if no
    /// workbook is bound — <see cref="ErrorService"/> falls back to the
    /// global slot on null.</summary>
    private static string? ActiveKey()
        => PyExcelServices.WorkbookContext.CurrentWorkbookKey;

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private void StubAction(IRibbonControl control, string name, string portNote)
    {
        _log.Info($"STUB onAction {name} (id={control.Id}) — port target: {portNote}");
    }
}
