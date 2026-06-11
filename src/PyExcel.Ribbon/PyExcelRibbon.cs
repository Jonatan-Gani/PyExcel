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

    /// <summary>Read the state for the currently-active workbook,
    /// returning <see cref="WorkbookState.Empty"/> if no workbook is
    /// active so every getter has a well-defined value to read.</summary>
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
    // getEnabled — single shared callback, mirroring v1 RibbonIsEnabled in
    // modRibbon.bas. Returns false until Phase 3 lands the StateService.
    // -------------------------------------------------------------------------

    public bool RibbonEnabled(IRibbonControl control) => ActiveState().Enabled;

    // -------------------------------------------------------------------------
    // Main group
    // -------------------------------------------------------------------------

    public void OnEnablePyExcel(IRibbonControl control)
    {
        // v1 (modRibbon.bas:461) ran a full setup wizard that provisioned
        // the workbook and then marked it enabled; that wizard is Phase 7.
        // For now this button is the enable/disable toggle for the active
        // workbook. Flipping Enabled fires StateChanged, which the
        // RibbonOnLoad handler turns into an IRibbonUI.Invalidate — so every
        // getEnabled-gated control lights up (or greys out) on the next
        // repaint without any extra wiring here.
        try
        {
            var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
            if (key is null) { _log.Info("OnEnablePyExcel: no active workbook"); return; }
            var now = !PyExcelServices.State.Get(key).Enabled;
            PyExcelServices.State.SetEnabled(key, now);
            _log.Info($"OnEnablePyExcel: workbook '{key}' enabled={now}");
        }
        catch (Exception ex)
        {
            _log.Error("OnEnablePyExcel failed", ex);
        }
    }

    public void OnSetup(IRibbonControl control)
    {
        _log.Info("OnSetup clicked");
        try
        {
            // An unsaved workbook has no location to anchor the environment to.
            var dir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
            if (string.IsNullOrEmpty(dir))
            {
                LogDisplay.WriteLine(
                    "Setup: save the workbook first — the Python environment is " +
                    "anchored to the workbook's location.");
                return;
            }
            // For a local workbook this is the workbook folder; for a
            // SharePoint/OneDrive-online workbook (whose folder is a URL) it
            // maps to a local %LOCALAPPDATA%\PyExcel folder. KernelHost resolves
            // the same directory at run time, so the kernel finds this venv.
            var projectDir = PyExcel.Common.ProjectDirectory.Resolve(dir);
            var success = SetupForm.Run(ExcelWindowOwner(), projectDir!, _log);
            if (success is not null)
                _log.Info($"OnSetup: finished, success={success}");
        }
        catch (Exception ex)
        {
            _log.Error("OnSetup failed", ex);
            LogDisplay.WriteLine($"Setup: {ex.Message}");
        }
    }

    public void OnOpenExplorer(IRibbonControl control)
    {
        _log.Info("OnOpenExplorer clicked");
        try
        {
            var dir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
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
            // If the active workbook's directory has a README.md, open
            // it with the user's default handler. Otherwise fall back to
            // an in-Excel alert pointing at the migration docs.
            var dir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
            if (!string.IsNullOrEmpty(dir))
            {
                var readme = Path.Combine(dir!, "README.md");
                if (File.Exists(readme))
                {
                    ShellLauncher.Open(readme);
                    return;
                }
            }

            const string text =
                "PyExcel v2.0 (alpha)\n\n" +
                "No README.md was found next to this workbook. " +
                "See the bundled docs/v2-build.md and ROADMAP.md for the " +
                "migration plan.";
            XlCall.Excel(XlCall.xlcAlert, text, 2 /* xlAlertWarning */);
        }
        catch (Exception ex)
        {
            _log.Error("OnReadMe failed", ex);
        }
    }

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
                    ProgressForm.StartModeless(ExcelWindowOwner(), "Running Python…"));
        }
        catch (Exception ex)
        {
            _log.Error("OnRunPython failed", ex);
        }
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
            var dir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
            if (string.IsNullOrEmpty(dir))
            {
                LogDisplay.WriteLine(
                    "Edit Python: the active workbook hasn't been saved yet — " +
                    "save the workbook first so the userScripts/ folder is " +
                    "located on disk.");
                return;
            }
            // Convention: scripts live under <workbookDir>/userScripts/<name>.
            // ScriptDirectoryWatcher uses the same root.
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

    public void OnScriptChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetSelectedScript(key, string.IsNullOrEmpty(text) ? null : text);
    }

    public string GetPyInput(IRibbonControl control) => ActiveState().PyInput ?? string.Empty;

    public void OnPyInputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetPyInput(key, text);
    }

    public string GetPyOutput(IRibbonControl control) => ActiveState().PyOutput ?? string.Empty;

    public void OnPyOutputChange(IRibbonControl control, string text)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) return;
        PyExcelServices.State.SetPyOutput(key, text);
    }

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
        PyExcelServices.State.SetSelectedAction(key, string.IsNullOrEmpty(text) ? null : text);
    }

    public void OnAddAction(IRibbonControl control)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) { _log.Info("OnAddAction: no active workbook"); return; }

        var state = PyExcelServices.State.Get(key);
        var result = EditActionForm.Prompt(
            ExcelWindowOwner(),
            state.AvailableScripts,
            ActionNames(state),
            existing: null,
            selectionProvider: CurrentSelectionAddress);
        if (result is null) { _log.Info("OnAddAction: cancelled"); return; }

        PyExcelServices.State.AddAction(key, result);
        _log.Info($"OnAddAction: saved '{result.Name}' to workbook '{key}'");
    }

    public void OnEditAction(IRibbonControl control)
    {
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        if (key is null) { _log.Info("OnEditAction: no active workbook"); return; }

        var state = PyExcelServices.State.Get(key);
        var existing = state.SelectedAction;
        if (existing is null) { _log.Info("OnEditAction: no action selected"); return; }

        var result = EditActionForm.Prompt(
            ExcelWindowOwner(),
            state.AvailableScripts,
            ActionNames(state),
            existing,
            selectionProvider: CurrentSelectionAddress);
        if (result is null) { _log.Info("OnEditAction: cancelled"); return; }

        // AddAction upserts by name, so a rename would leave the original
        // entry behind under its old name — drop it first when the name
        // changed. (The validator already guaranteed the new name doesn't
        // collide with a *different* action.)
        if (!string.Equals(existing.Name, result.Name, StringComparison.Ordinal))
            PyExcelServices.State.DeleteAction(key, existing.Name);
        PyExcelServices.State.AddAction(key, result);
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

    /// <summary>The active selection's address as <c>Sheet!A1:B2</c>, for
    /// the range picker's "Use current selection" button. Returns null when
    /// the selection isn't a range (a chart, a shape) or COM is unhappy —
    /// the picker then just hides the button / keeps the typed text.</summary>
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
                CurrentSelectionAddress);
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
                CurrentSelectionAddress);
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
                ExcelWindowOwner(), state.PasteOutput, CurrentSelectionAddress);
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

    /// <summary>Show the last error in Excel-DNA's <see cref="LogDisplay"/>.
    /// Brings the window to front so the user sees it even if it was
    /// previously dismissed.</summary>
    public void OnShowLastError(IRibbonControl control)
    {
        try
        {
            var record = PyExcelServices.Errors.GetLast(ActiveKey());
            if (record is null)
            {
                LogDisplay.WriteLine("[PyExcel] No errors recorded.");
            }
            else
            {
                LogDisplay.WriteLine(record.FormatForClipboard());
            }
            LogDisplay.Show();
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
