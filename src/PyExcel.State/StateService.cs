using System;
using System.Collections.Generic;
using System.Linq;

namespace PyExcel.State;

/// <summary>
/// In-memory registry of each workbook's PyExcel project, the single source of
/// truth for the ribbon's getEnabled / get* callbacks.
///
/// <para><b>Per sheet.</b> A workbook's configuration is split: the
/// <em>workbook-scoped</em> facts (<see cref="WorkbookState.Enabled"/>,
/// <see cref="WorkbookState.ProjectDir"/>, the shared
/// <see cref="WorkbookState.AvailableScripts"/>) plus a map of
/// <em>sheet-scoped</em> <see cref="SheetProfile"/>s (the selected script, the
/// Run/Import/Export/Paste field bindings, the saved actions). Each workbook
/// tracks a <em>current sheet</em>; <see cref="Get"/> returns a
/// <see cref="WorkbookState"/> that projects the current sheet's profile onto the
/// flat fields, so the ribbon and the data services read the active sheet's
/// configuration without knowing sheets exist. <see cref="SetCurrentSheet"/>
/// (driven by the COM <c>SheetActivate</c> / <c>WorkbookActivate</c> handlers)
/// switches which sheet is projected.</para>
///
/// <para>Thread-safe by a single coarse lock — mutations are infrequent (user
/// actions, file-system events, COM workbook events). <see cref="StateChanged"/>
/// fires synchronously after each mutation; subscribers must not call back into
/// the service (the ribbon queues <c>IRibbonUI.Invalidate</c> instead).</para>
///
/// <para>Persistence lives elsewhere: <see cref="GetProfile"/> snapshots a
/// workbook's full <see cref="WorkbookProfileData"/> (all sheets) for the
/// CustomXMLPart codec, and <see cref="LoadProfile"/> replays it on restore.</para>
/// </summary>
public sealed class StateService
{
    /// <summary>Mutable per-workbook store: workbook-scoped facts plus the
    /// sheet-scoped profile map and a current-sheet pointer.</summary>
    private sealed class Entry
    {
        public bool Enabled;
        public string? ProjectDir;
        public IReadOnlyList<string> AvailableScripts = Array.Empty<string>();
        public string CurrentSheet = WorkbookProfileData.DefaultSheetKey;
        public readonly Dictionary<string, SheetProfile> Sheets = new(StringComparer.Ordinal);

        /// <summary>The profile in effect for <paramref name="sheet"/>: its own
        /// entry, else the inherited default-bucket entry, else empty.</summary>
        public SheetProfile Effective(string sheet)
            => Sheets.TryGetValue(sheet, out var p) ? p
               : Sheets.TryGetValue(WorkbookProfileData.DefaultSheetKey, out var d) ? d
               : SheetProfile.Empty;
    }

    private readonly Dictionary<string, Entry> _entries = new(StringComparer.Ordinal);
    private readonly object _lock = new();

    /// <summary>Fired after a successful mutation. Carries the affected workbook
    /// key so subscribers can skip work if they only care about the active one.</summary>
    public event EventHandler<StateChangedEventArgs>? StateChanged;

    /// <summary>Read the current state for a workbook, projecting its current
    /// sheet's profile onto the flat fields. Returns
    /// <see cref="WorkbookState.Empty"/> for a workbook never touched — callers
    /// never need a null-check.</summary>
    public WorkbookState Get(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        lock (_lock)
        {
            return _entries.TryGetValue(workbookKey, out var e)
                ? Project(workbookKey, e)
                : WorkbookState.Empty(workbookKey);
        }
    }

    /// <summary>Atomic read-modify-write of the <em>current sheet's</em> view. The
    /// <paramref name="mutator"/> receives the projected <see cref="WorkbookState"/>
    /// and returns the new one; workbook-scoped fields are written back to the
    /// workbook, and the sheet-scoped fields to the current sheet's profile (only
    /// materialising a sheet entry when the sheet fields actually changed, so a
    /// workbook-scoped edit never accidentally copies an inherited default).</summary>
    public void Update(string workbookKey, Func<WorkbookState, WorkbookState> mutator)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        if (mutator is null) throw new ArgumentNullException(nameof(mutator));

        lock (_lock)
        {
            _entries.TryGetValue(workbookKey, out var existing);
            var current = existing is null
                ? WorkbookState.Empty(workbookKey)
                : Project(workbookKey, existing);
            var next = mutator(current);
            if (next is null)
                throw new InvalidOperationException(
                    "mutator returned null; return WorkbookState.Empty(key) instead");
            if (!string.Equals(next.WorkbookKey, workbookKey, StringComparison.Ordinal))
                throw new InvalidOperationException(
                    $"mutator changed WorkbookKey from '{workbookKey}' to '{next.WorkbookKey}'");

            // Create the entry only now — a throwing/invalid mutator above leaves
            // the registry untouched.
            var e = existing ?? new Entry();

            // Workbook-scoped fields: apply unconditionally (idempotent for a
            // sheet-scoped edit, which leaves them unchanged).
            e.Enabled = next.Enabled;
            e.ProjectDir = next.ProjectDir;
            e.AvailableScripts = next.AvailableScripts;

            // Sheet-scoped fields: write to the current sheet, but only
            // materialise an entry when they actually changed from what the sheet
            // currently shows (its own entry, or the inherited default). This is
            // what stops a workbook-scoped edit (e.g. SetAvailableScripts) from
            // copying an inherited default into the active sheet.
            var sheet = e.CurrentSheet;
            var nextSheet = ToProfile(next);
            if (e.Sheets.ContainsKey(sheet) || !SameSheet(nextSheet, e.Effective(sheet)))
                e.Sheets[sheet] = nextSheet;

            if (existing is null) _entries[workbookKey] = e;
        }
        StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    /// <summary>Switch which sheet <see cref="Get"/> projects for a workbook.
    /// Driven by the COM <c>SheetActivate</c> / <c>WorkbookActivate</c> handlers.
    /// A null/blank name selects the workbook's default bucket. Fires
    /// <see cref="StateChanged"/> only when the active sheet actually changed, so
    /// re-activating the same sheet doesn't trigger a redundant repaint.</summary>
    public void SetCurrentSheet(string workbookKey, string? sheet)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        var normalized = string.IsNullOrEmpty(sheet) ? WorkbookProfileData.DefaultSheetKey : sheet!;
        bool changed;
        lock (_lock)
        {
            var e = GetOrCreate(workbookKey);
            changed = !string.Equals(e.CurrentSheet, normalized, StringComparison.Ordinal);
            e.CurrentSheet = normalized;
        }
        if (changed) StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    /// <summary>Forget a workbook — called on <c>Workbook.BeforeClose</c> so closed
    /// workbooks don't accumulate forever in memory.</summary>
    public void Forget(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        bool removed;
        lock (_lock) { removed = _entries.Remove(workbookKey); }
        if (removed) StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    /// <summary>Snapshot of all currently-tracked workbook keys. Test helper.</summary>
    public IReadOnlyList<string> KnownWorkbooks()
    {
        lock (_lock) { return _entries.Keys.ToList(); }
    }

    // -------------------------------------------------------------------------
    // Persistence bridge — the full multi-sheet structure for the codec.
    // -------------------------------------------------------------------------

    /// <summary>Snapshot the workbook's full <see cref="WorkbookProfileData"/>
    /// (workbook-scoped fields + every configured sheet) for persistence. Returns
    /// <see cref="WorkbookProfileData.Empty"/> for an untracked workbook.</summary>
    public WorkbookProfileData GetProfile(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        lock (_lock)
        {
            if (!_entries.TryGetValue(workbookKey, out var e)) return WorkbookProfileData.Empty;
            var sheets = new Dictionary<string, SheetProfile>(StringComparer.Ordinal);
            foreach (var kv in e.Sheets)
                if (kv.Value.IsConfigured) sheets[kv.Key] = kv.Value;
            return new WorkbookProfileData
            {
                Enabled = e.Enabled,
                ProjectDir = e.ProjectDir,
                Sheets = sheets,
            };
        }
    }

    /// <summary>Replace a workbook's workbook-scoped fields and sheet map from a
    /// restored <see cref="WorkbookProfileData"/>. The transient
    /// <see cref="WorkbookState.AvailableScripts"/> and the current-sheet pointer
    /// are preserved (the caller repopulates scripts from the live userScripts
    /// folder and points the pointer at the live active sheet).</summary>
    public void LoadProfile(string workbookKey, WorkbookProfileData data)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        if (data is null) throw new ArgumentNullException(nameof(data));
        lock (_lock)
        {
            var e = GetOrCreate(workbookKey);
            e.Enabled = data.Enabled;
            e.ProjectDir = data.ProjectDir;
            e.Sheets.Clear();
            foreach (var kv in data.Sheets)
                e.Sheets[kv.Key] = kv.Value;
        }
        StateChanged?.Invoke(this, new StateChangedEventArgs(workbookKey));
    }

    // -------------------------------------------------------------------------
    // Projection helpers
    // -------------------------------------------------------------------------

    private Entry GetOrCreate(string key)
    {
        if (!_entries.TryGetValue(key, out var e))
        {
            e = new Entry();
            _entries[key] = e;
        }
        return e;
    }

    private static WorkbookState Project(string key, Entry e)
    {
        var sheet = e.CurrentSheet;
        var p = e.Effective(sheet);
        return WorkbookState.Empty(key) with
        {
            Enabled = e.Enabled,
            ProjectDir = e.ProjectDir,
            CurrentSheet = sheet.Length == 0 ? null : sheet,
            AvailableScripts = e.AvailableScripts,
            SelectedScript = p.SelectedScript,
            PyInput = p.PyInput,
            PyOutput = p.PyOutput,
            Actions = p.Actions,
            SelectedActionName = p.SelectedActionName,
            ImportInput = p.ImportInput,
            ImportOutput = p.ImportOutput,
            ExportInput = p.ExportInput,
            ExportOutput = p.ExportOutput,
            PasteOutput = p.PasteOutput,
        };
    }

    private static SheetProfile ToProfile(WorkbookState s) => new()
    {
        SelectedScript = s.SelectedScript,
        PyInput = s.PyInput,
        PyOutput = s.PyOutput,
        Actions = s.Actions,
        SelectedActionName = s.SelectedActionName,
        ImportInput = s.ImportInput,
        ImportOutput = s.ImportOutput,
        ExportInput = s.ExportInput,
        ExportOutput = s.ExportOutput,
        PasteOutput = s.PasteOutput,
    };

    /// <summary>Whether two sheet profiles carry the same sheet-scoped data.
    /// Strings compare by value; <see cref="SheetProfile.Actions"/> compares by
    /// reference — a workbook-scoped edit passes the projection's list through
    /// unchanged (same reference), whereas a sheet-scoped action edit rebuilds it,
    /// so this exactly distinguishes "did the sheet change".</summary>
    private static bool SameSheet(SheetProfile a, SheetProfile b)
        => ReferenceEquals(a.Actions, b.Actions)
           && a.SelectedScript == b.SelectedScript
           && a.PyInput == b.PyInput
           && a.PyOutput == b.PyOutput
           && a.SelectedActionName == b.SelectedActionName
           && a.ImportInput == b.ImportInput
           && a.ImportOutput == b.ImportOutput
           && a.ExportInput == b.ExportInput
           && a.ExportOutput == b.ExportOutput
           && a.PasteOutput == b.PasteOutput;

    // -------------------------------------------------------------------------
    // Typed helpers — every ribbon edit-event funnels through Update via one of
    // these. Workbook-scoped (Enabled / ProjectDir / AvailableScripts) and
    // sheet-scoped (everything else) helpers look identical here; Update routes
    // each to the right place. Keeping them in one place means a future "log
    // every state change" hook needs to instrument only this class.
    // -------------------------------------------------------------------------

    public void SetEnabled(string key, bool enabled)
        => Update(key, s => s with { Enabled = enabled });

    public void SetAvailableScripts(string key, IReadOnlyList<string> scripts)
        => Update(key, s => s with { AvailableScripts = scripts });

    public void SetProjectDir(string key, string? dir)
        => Update(key, s => s with { ProjectDir = dir });

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
            // Replace if name already exists — Add is also the upsert path the
            // Edit form uses.
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
    /// output become the current sheet's <see cref="WorkbookState.SelectedScript"/>,
    /// <see cref="WorkbookState.PyInput"/>, and <see cref="WorkbookState.PyOutput"/>
    /// — and select it, in one atomic update (so the ribbon invalidates once).</summary>
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
}

/// <summary>Carries the affected workbook key. Subscribers can skip work cheaply
/// by comparing against the active key.</summary>
public sealed class StateChangedEventArgs : EventArgs
{
    public string WorkbookKey { get; }
    public StateChangedEventArgs(string workbookKey)
    {
        WorkbookKey = workbookKey ?? throw new ArgumentNullException(nameof(workbookKey));
    }
}
