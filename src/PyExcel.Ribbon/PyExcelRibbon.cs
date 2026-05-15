using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using PyExcel.Common.Logging;

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
/// is added (Phase 4), the OnRunPython callback will enqueue work to a
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
    }

    public object? LoadImage(string imageName)
    {
        // imageName="customLogo" — load the embedded PNG. Phase 1 returns
        // null so Excel falls back to no image; Phase 6 will ship the PNG
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

    public bool RibbonEnabled(IRibbonControl control)
    {
        // PHASE 3: read from StateService — has the active workbook been
        // Enabled? For now: false everywhere except the bootstrap buttons
        // (which the XML omits from getEnabled wiring).
        return false;
    }

    // -------------------------------------------------------------------------
    // Main group
    // -------------------------------------------------------------------------

    public void OnEnablePyExcel(IRibbonControl control)
        => StubAction(control, "OnEnablePyExcel", "modRibbon.bas:461 — invokes setup wizard");

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
            // ported yet; that lives in PyExcel.Common.Shell in Phase 7.
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
        => StubAction(control, "OnRunPython",
            "modRibbon.bas:1466 — orchestrates: read script, marshal input range, " +
            "send Run frame, write output. SAFE-1: this must enqueue and return; " +
            "never block on the pipe.");

    public void OnEditPython(IRibbonControl control)
        => StubAction(control, "OnEditPython", "modRibbon.bas:1400 — shells the user's editor");

    public int GetScriptCount(IRibbonControl control) => 0;
    public string GetScriptLabel(IRibbonControl control, int index) => string.Empty;
    public string GetScriptText(IRibbonControl control) => string.Empty;
    public void OnScriptChange(IRibbonControl control, string text)
        => StubChange(control, "OnScriptChange", "modRibbon.bas:1159", text);

    public string GetPyInput(IRibbonControl control) => string.Empty;
    public void OnPyInputChange(IRibbonControl control, string text)
        => StubChange(control, "OnPyInputChange", "modRibbon.bas:980", text);

    public string GetPyOutput(IRibbonControl control) => string.Empty;
    public void OnPyOutputChange(IRibbonControl control, string text)
        => StubChange(control, "OnPyOutputChange", "modRibbon.bas:1023", text);

    public int GetActionCount(IRibbonControl control) => 0;
    public string GetActionLabel(IRibbonControl control, int index) => string.Empty;
    public string GetActionText(IRibbonControl control) => string.Empty;
    public void OnActionChange(IRibbonControl control, string text)
        => StubChange(control, "OnActionChange", "modRibbon.bas:1683", text);

    public void OnAddAction(IRibbonControl control)
        => StubAction(control, "OnAddAction", "modRibbon.bas:1340 — shows EditActionForm");

    public void OnEditAction(IRibbonControl control)
        => StubAction(control, "OnEditAction", "modRibbon.bas:1447 — shows EditActionForm pre-populated");

    public void OnDeleteAction(IRibbonControl control)
        => StubAction(control, "OnDeleteAction", "modRibbon.bas:1736 — removes action from state");

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
