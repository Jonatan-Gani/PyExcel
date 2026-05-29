using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using PyExcel.Common.Logging;
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
        // imageName="customLogo" — load the embedded PNG. Phase 1 returns
        // null so Excel falls back to no image; Phase 8 will ship the PNG
        // as an EmbeddedResource and return a System.Drawing.Bitmap here.
        if (string.Equals(imageName, "customLogo", StringComparison.Ordinal))
        {
            return null;
        }
        return null;
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

    public void OnOpenExplorer(IRibbonControl control)
        => StubAction(control, "OnOpenExplorer", "modRibbon.bas:515 — shells explorer.exe at project root");

    public void OnReadMe(IRibbonControl control)
    {
        _log.Info("OnReadMe clicked");
        try
        {
            // Phase 1 deliverable: show the readme. We pop a MessageBox
            // rather than launching the user's text editor because the
            // file-association resolution requires Shell APIs we haven't
            // ported yet; that lives in PyExcel.Common.Shell in Phase 5.
            const string text =
                "PyExcel v2.0 (alpha)\n\n" +
                "This is the .NET rewrite of the PyExcel add-in. Phase 1 — " +
                "ribbon skeleton — is the only thing wired up today.\n\n" +
                "See README.md and docs/v2-build.md for the migration plan.";

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
            PyExcel.Excel.RangeRunner.RunActiveScript(state);
        }
        catch (Exception ex)
        {
            _log.Error("OnRunPython failed", ex);
        }
    }

    public void OnEditPython(IRibbonControl control)
        => StubAction(control, "OnEditPython", "modRibbon.bas:1400 — shells the user's editor");

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
        => StubAction(control, "OnAddAction",
            "modRibbon.bas:1340 — shows EditActionForm. " +
            "Phase 8 ships the form; Phase 3 wires StateService.AddAction once it returns.");

    public void OnEditAction(IRibbonControl control)
        => StubAction(control, "OnEditAction",
            "modRibbon.bas:1447 — shows EditActionForm pre-populated. " +
            "Phase 8 ships the form; Phase 3 wires StateService.AddAction (upserts) once it returns.");

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
    // Import group
    // -------------------------------------------------------------------------

    public void OnImport(IRibbonControl control)
        => StubAction(control, "OnImport", "modRibbon.bas:593 — runs ImportService");

    public string GetImportInput(IRibbonControl control) => string.Empty;
    public void OnImportInputChange(IRibbonControl control, string text)
        => StubChange(control, "OnImportInputChange", "modRibbon.bas:637", text);

    public string GetImportOutput(IRibbonControl control) => string.Empty;
    public void OnImportOutputChange(IRibbonControl control, string text)
        => StubChange(control, "OnImportOutputChange", "modRibbon.bas:681", text);

    public void OnEditImport(IRibbonControl control)
        => StubAction(control, "OnEditImport", "modRibbon.bas:708 — shows EditImportForm");

    // -------------------------------------------------------------------------
    // Export group
    // -------------------------------------------------------------------------

    public void OnExport(IRibbonControl control)
        => StubAction(control, "OnExport", "modRibbon.bas:740 — runs ExportService");

    public string GetExportInput(IRibbonControl control) => string.Empty;
    public void OnExportInputChange(IRibbonControl control, string text)
        => StubChange(control, "OnExportInputChange", "modRibbon.bas:784", text);

    public string GetExportOutput(IRibbonControl control) => string.Empty;
    public void OnExportOutputChange(IRibbonControl control, string text)
        => StubChange(control, "OnExportOutputChange", "modRibbon.bas:827", text);

    public void OnEditExport(IRibbonControl control)
        => StubAction(control, "OnEditExport", "modRibbon.bas:853 — shows EditExportForm");

    // -------------------------------------------------------------------------
    // Paste group
    // -------------------------------------------------------------------------

    public void OnPaste(IRibbonControl control)
        => StubAction(control, "OnPaste", "modRibbon.bas:866 — pastes saved artifact");

    public string GetPasteOutput(IRibbonControl control) => string.Empty;
    public void OnPasteOutputChange(IRibbonControl control, string text)
        => StubChange(control, "OnPasteOutputChange", "modRibbon.bas:913", text);

    public void OnEditPaste(IRibbonControl control)
        => StubAction(control, "OnEditPaste", "modRibbon.bas:940 — shows EditPasteForm");

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private void StubAction(IRibbonControl control, string name, string portNote)
    {
        _log.Info($"STUB onAction {name} (id={control.Id}) — port target: {portNote}");
    }

    private void StubChange(IRibbonControl control, string name, string portRef, string text)
    {
        _log.Debug($"STUB onChange {name} (id={control.Id}, text='{text}') — port target: {portRef}");
    }
}
