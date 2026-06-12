using System;
using System.Collections.Generic;
using System.Linq;

namespace PyExcel.State;

/// <summary>
/// In-memory registry of <see cref="WorkbookState"/> per workbook key.
/// Single source of truth for the ribbon's getEnabled / get* callbacks.
///
/// <para>Thread-safe by a single coarse lock — mutations are infrequent
/// (user actions, file-system events, COM workbook events) so the lock
/// is not a hot path. Reads return the immutable
/// <see cref="WorkbookState"/> record, which the caller can keep walking
/// after the lock is released without fear of torn writes.</para>
///
/// <para><see cref="StateChanged"/> fires synchronously after each
/// mutation. Subscribers must not call back into the service (which
/// would re-enter the lock and deadlock); the ribbon's handler queues
/// the <c>IRibbonUI.Invalidate</c> call to the macro queue instead of
/// invoking it inline.</para>
///
/// <para>This service does not own persistence. The Phase 3
/// <c>CustomXMLPart</c> serializer is a separate component that
/// snapshots the dictionary on <c>Workbook.BeforeSave</c> and replays
/// it on <c>Workbook.Open</c>.</para>
/// </summary>
public sealed class StateService
{
    private readonly Dictionary<string, WorkbookState> _states = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <summary>Fired after a successful <see cref="Update"/> (or any
    /// helper). Carries the affected workbook key so subscribers can
    /// skip work if they only care about the active workbook.</summary>
    public event EventHandler<StateChangedEventArgs>? StateChanged;

    /// <summary>Read the current state for a workbook. Returns
    /// <see cref="WorkbookState.Empty"/> if the workbook hasn't been
    /// touched yet — callers never need a null-check.</summary>
    public WorkbookState Get(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        lock (_lock)
        {
            return _states.TryGetValue(workbookKey, out var s) ? s : WorkbookState.Empty(workbookKey);
        }
    }

    /// <summary>Atomic read-modify-write. The <paramref name="mutator"/>
    /// receives the current state and returns the new one; the service
    /// stores the result and fires <see cref="StateChanged"/>.</summary>
    public void Update(string workbookKey, Func<WorkbookState, WorkbookState> mutator)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        if (mutator is null) throw new ArgumentNullException(nameof(mutator));

        WorkbookState next;
        lock (_lock)
        {
            var current = _states.TryGetValue(workbookKey, out var s)
                ? s
                : WorkbookState.Empty(workbookKey);
            next = mutator(current);
            if (next is null)
                throw new InvalidOperationException(
                    "mutator returned null; return WorkbookState.Empty(key) instead");
            if (!string.Equals(next.WorkbookKey, workbookKey, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    $"mutator changed WorkbookKey from '{workbookKey}' to '{next.WorkbookKey}'");
            _states[workbookKey] = next;
        }
        StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    /// <summary>Forget a workbook — called on <c>Workbook.BeforeClose</c>
    /// so closed workbooks don't accumulate forever in memory.</summary>
    public void Forget(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        bool removed;
        lock (_lock) { removed = _states.Remove(workbookKey); }
        if (removed) StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    /// <summary>Snapshot of all currently-tracked workbook keys.
    /// Test helper; not used at runtime.</summary>
    public IReadOnlyList<string> KnownWorkbooks()
    {
        lock (_lock) { return _states.Keys.ToList(); }
    }

    // -------------------------------------------------------------------------
    // Typed helpers — every ribbon edit-event funnels through Update via one
    // of these. Keeping them in one place means a future "log every state
    // change" hook needs to instrument only this class.
    // -------------------------------------------------------------------------

    public void SetEnabled(string key, bool enabled)
        => Update(key, s => s with { Enabled = enabled });

    public void SetCurrentSheet(string key, string? sheet)
        => Update(key, s => s with { CurrentSheet = sheet });

    public void SetAvailableScripts(string key, IReadOnlyList<string> scripts)
        => Update(key, s => s with { AvailableScripts = scripts });

    public void SetSelectedScript(string key, string? script)
        => Update(key, s => s with { SelectedScript = script });

    public void SetPyInput(string key, string? value)
        => Update(key, s => s with { PyInput = value });

    public void SetPyOutput(string key, string? value)
        => Update(key, s => s with { PyOutput = value });

    public void AddAction(string key, RibbonAction action)
    {
        if (action is null) throw new ArgumentNullException(nameof(action));
        Update(key, s =>
        {
            var list = s.Actions.ToList();
            // Replace if name already exists — Add is also the upsert path
            // the Edit form uses.
            var existing = list.FindIndex(a =>
                string.Equals(a.Name, action.Name, StringComparison.Ordinal));
            if (existing >= 0) list[existing] = action;
            else list.Add(action);
            return s with { Actions = list, SelectedActionName = action.Name };
        });
    }

    public void DeleteAction(string key, string actionName)
    {
        if (actionName is null) throw new ArgumentNullException(nameof(actionName));
        Update(key, s =>
        {
            var list = s.Actions.Where(a =>
                !string.Equals(a.Name, actionName, StringComparison.Ordinal)).ToList();
            var selected = string.Equals(s.SelectedActionName, actionName, StringComparison.Ordinal)
                ? null
                : s.SelectedActionName;
            return s with { Actions = list, SelectedActionName = selected };
        });
    }

    public void SetSelectedAction(string key, string? name)
        => Update(key, s => s with { SelectedActionName = name });

    /// <summary>Load a saved action into the run boxes — its script, input, and
    /// output become <see cref="WorkbookState.SelectedScript"/>,
    /// <see cref="WorkbookState.PyInput"/>, and <see cref="WorkbookState.PyOutput"/>
    /// — and select it, in one atomic update (so the ribbon invalidates once).
    /// The Run button reads those boxes, so this is what makes picking or saving
    /// an action actually runnable.</summary>
    public void LoadAction(string key, RibbonAction action)
    {
        if (action is null) throw new ArgumentNullException(nameof(action));
        Update(key, s => s with
        {
            SelectedScript = action.Script,
            PyInput = action.Input,
            PyOutput = action.Output,
            SelectedActionName = action.Name,
        });
    }

    // -------------------------------------------------------------------------
    // Import / Export / Paste — text fields owned by Phase 5's data
    // services. The COM-bound services (ImportService / ExportService /
    // PasteService) read these on click; the ribbon's onChange callbacks
    // persist edits here. See the ribbon's Import / Export / Paste groups.
    // -------------------------------------------------------------------------

    public void SetImportInput(string key, string? value)
        => Update(key, s => s with { ImportInput = value });

    public void SetImportOutput(string key, string? value)
        => Update(key, s => s with { ImportOutput = value });

    public void SetExportInput(string key, string? value)
        => Update(key, s => s with { ExportInput = value });

    public void SetExportOutput(string key, string? value)
        => Update(key, s => s with { ExportOutput = value });

    public void SetPasteOutput(string key, string? value)
        => Update(key, s => s with { PasteOutput = value });

    /// <summary>The dedicated project directory the user chose for this
    /// workbook on Enable (where Setup provisions the venv/kernel/userScripts).
    /// Persisted by the codec; the runtime kernel and the ribbon's
    /// userScripts lookup prefer it over the workbook-derived default.</summary>
    public void SetProjectDir(string key, string? dir)
        => Update(key, s => s with { ProjectDir = dir });
}

/// <summary>Carries the affected workbook key. Subscribers can skip
/// work cheaply by comparing against the active key.</summary>
public sealed class StateChangedEventArgs : EventArgs
{
    public string WorkbookKey { get; }
    public StateChangedEventArgs(string workbookKey)
    {
        WorkbookKey = workbookKey ?? throw new ArgumentNullException(nameof(workbookKey));
    }
}
