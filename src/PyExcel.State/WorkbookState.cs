using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// Immutable per-workbook state — the projection the ribbon reads to
/// render every getter. New states are produced by
/// <see cref="StateService.Update"/>; the previous instance keeps living
/// as long as someone holds a reference, so a ribbon callback reading
/// in-flight cannot see a torn write.
/// </summary>
/// <remarks>
/// The state intentionally carries only what the ribbon needs to render.
/// Pipe/kernel ownership is in <see cref="PyExcel.Excel.KernelHost"/>,
/// not here; Phase 3 deliberately keeps state and runtime decoupled so
/// neither side has to know the other's lifecycle.
/// </remarks>
public sealed record WorkbookState(
    string WorkbookKey,
    bool Enabled,
    string? CurrentSheet,
    IReadOnlyList<string> AvailableScripts,
    string? SelectedScript,
    string? PyInput,
    string? PyOutput,
    IReadOnlyList<RibbonAction> Actions,
    string? SelectedActionName,
    string? ImportInput = null,
    string? ImportOutput = null,
    string? ExportInput = null,
    string? ExportOutput = null,
    string? PasteOutput = null,
    string? ProjectDir = null)
{
    /// <summary>The all-defaults state used when a workbook is seen for
    /// the first time, before any user action.</summary>
    public static WorkbookState Empty(string workbookKey) => new(
        WorkbookKey: workbookKey,
        Enabled: false,
        CurrentSheet: null,
        AvailableScripts: Array.Empty<string>(),
        SelectedScript: null,
        PyInput: null,
        PyOutput: null,
        Actions: Array.Empty<RibbonAction>(),
        SelectedActionName: null,
        ImportInput: null,
        ImportOutput: null,
        ExportInput: null,
        ExportOutput: null,
        PasteOutput: null,
        ProjectDir: null);

    /// <summary>The currently-selected <see cref="RibbonAction"/>, or
    /// <see langword="null"/> if no action is selected (or the selection
    /// doesn't resolve).</summary>
    public RibbonAction? SelectedAction
    {
        get
        {
            if (SelectedActionName is null) return null;
            foreach (var a in Actions)
                if (string.Equals(a.Name, SelectedActionName, StringComparison.Ordinal))
                    return a;
            return null;
        }
    }
}

/// <summary>
/// One saved Python action — script + input + output + optional kwargs —
/// configured by the user via the Add/Edit form. The form itself lands
/// in Phase 8; Phase 3 only owns the in-memory model and persistence.
/// </summary>
public sealed record RibbonAction(
    string Name,
    string Script,
    string Input,
    string Output,
    IReadOnlyDictionary<string, string>? Kwargs = null);
