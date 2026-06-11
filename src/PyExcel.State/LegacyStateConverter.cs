using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// The v1 → v2 per-workbook state migration (Phase 9, Decision #3).
///
/// <para>v1 stored its per-<em>sheet</em> state in Excel defined Names
/// (<c>Actions</c>, <c>cmbScript</c>, <c>txtPyInput</c>, …); v2 stores one
/// per-<em>workbook</em> <see cref="WorkbookState"/> as a
/// <c>CustomXMLPart</c> (see <see cref="WorkbookStateCodec"/>). This class is
/// the pure, COM-free core of the converter: given the raw string values v1
/// wrote into those Names <em>for one sheet</em>, it produces the equivalent
/// <see cref="WorkbookState"/>. The Windows-only driver reads the Names off a
/// live workbook (via <c>Name.RefersTo</c> / <c>Evaluate</c>, exactly as v1's
/// <c>GetSheetValue</c> did) and writes the result through the codec — that COM
/// tail is the Windows smoke test, same split as every other phase.</para>
///
/// <para><b>Per-sheet → per-workbook.</b> v1 kept independent state on every
/// sheet; v2 has a single workbook-scoped state. A migration therefore picks
/// one sheet's worth of v1 Names to carry forward (the driver chooses — the v1
/// "active" sheet, or the first sheet that has any state). Collapsing several
/// sheets is out of scope here by design: it has no lossless mapping, so the
/// decision stays with the caller rather than being baked into the parser.</para>
///
/// <para>What is <em>not</em> carried: v1's per-action <c>entireRow</c> /
/// <c>entreToEnd</c> flags (the v2 architecture designed row-spill out — see
/// the roadmap's "Designed out" list), and v1's per-sheet scoping. v1 had no
/// kwargs, so migrated actions always have <see cref="RibbonAction.Kwargs"/>
/// null.</para>
/// </summary>
public static class LegacyStateConverter
{
    /// <summary>The exact v1 defined-Name keys, kept here as the single
    /// source of truth so the COM-side driver and these docs can't drift from
    /// what v1's <c>modRibbon.bas</c> actually wrote.</summary>
    public static class LegacyNames
    {
        /// <summary>Workbook-scoped. <c>"1"</c> / <c>"0"</c>.</summary>
        public const string Enabled = "PyExcelEnabled";

        /// <summary>Sheet-scoped. The selected action's name.</summary>
        public const string SelectedAction = "SelectedAction";

        /// <summary>Sheet-scoped. The serialized action list — see
        /// <see cref="ParseActions"/> for the format.</summary>
        public const string Actions = "Actions";

        /// <summary>Sheet-scoped. The selected script file name.</summary>
        public const string SelectedScript = "cmbScript";

        /// <summary>Sheet-scoped. The Run-Python input range(s).</summary>
        public const string PyInput = "txtPyInput";

        /// <summary>Sheet-scoped. The Run-Python output range.</summary>
        public const string PyOutput = "txtPyOutput";

        /// <summary>Sheet-scoped. The import source path.</summary>
        public const string ImportInput = "txtImportInput";

        /// <summary>Sheet-scoped. The import destination range.</summary>
        public const string ImportOutput = "txtImportOutput";

        /// <summary>Sheet-scoped. The export source range.</summary>
        public const string ExportInput = "txtExportInput";

        /// <summary>Sheet-scoped. The export destination path.</summary>
        public const string ExportOutput = "txtExportOutput";

        /// <summary>Sheet-scoped. The paste destination range.</summary>
        public const string PasteOutput = "txtPasteOutput";
    }

    /// <summary>The v1 action separator (ASCII Group Separator, <c>Chr(29)</c>)
    /// — chosen by v1 because it never appears in a range ref or script name.</summary>
    private const char GroupSeparator = '\u001D';

    /// <summary>Convert the raw v1 Name values captured for one sheet
    /// (<paramref name="legacy"/>) into a fresh <see cref="WorkbookState"/>
    /// keyed by <paramref name="workbookKey"/>. Blank text fields become
    /// <see langword="null"/> (so the codec omits them, matching "never set");
    /// every value is trimmed. Transient fields
    /// (<see cref="WorkbookState.CurrentSheet"/>,
    /// <see cref="WorkbookState.AvailableScripts"/>) are left at their defaults
    /// for the caller to repopulate from live sources.</summary>
    public static WorkbookState Convert(LegacyWorkbookState legacy, string workbookKey)
    {
        if (legacy is null) throw new ArgumentNullException(nameof(legacy));
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));

        return WorkbookState.Empty(workbookKey) with
        {
            Enabled = ParseEnabled(legacy.Enabled),
            SelectedScript = Optional(legacy.SelectedScript),
            PyInput = Optional(legacy.PyInput),
            PyOutput = Optional(legacy.PyOutput),
            SelectedActionName = Optional(legacy.SelectedAction),
            Actions = ParseActions(legacy.Actions),
            ImportInput = Optional(legacy.ImportInput),
            ImportOutput = Optional(legacy.ImportOutput),
            ExportInput = Optional(legacy.ExportInput),
            ExportOutput = Optional(legacy.ExportOutput),
            PasteOutput = Optional(legacy.PasteOutput),
        };
    }

    /// <summary>Parse v1's serialized <c>Actions</c> Name value into the v2
    /// action list. A direct port of v1's <c>LoadActionsForSheet</c>:
    /// <list type="bullet">
    ///   <item>Actions are separated by <c>Chr(29)</c>; older data falls back
    ///     to <c>Chr(10)</c> then <c>";"</c> (the same auto-detection v1
    ///     used to read data written by earlier builds).</item>
    ///   <item>Each action is pipe-delimited. The first token is the name.</item>
    ///   <item>If the second token contains <c>=</c> it's the named format
    ///     (<c>key=value</c> fields, where <c>input</c>/<c>output</c> may repeat
    ///     and accumulate with <c>"; "</c>); otherwise it's the legacy
    ///     positional format (<c>name|script|input|output|…</c>).</item>
    /// </list>
    /// The v1-only <c>entireRow</c> / <c>entreToEnd</c> fields are dropped;
    /// duplicate names keep the first occurrence (v2 treats action names as a
    /// unique upsert key, mirroring v1's dictionary-keyed storage).</summary>
    public static IReadOnlyList<RibbonAction> ParseActions(string? raw)
    {
        if (string.IsNullOrEmpty(raw)) return Array.Empty<RibbonAction>();

        // Separator auto-detection, newest format first (matches v1).
        string rowSep =
            raw!.IndexOf(GroupSeparator) >= 0 ? GroupSeparator.ToString()
            : raw.IndexOf('\n') >= 0 ? "\n"
            : ";";

        var actions = new List<RibbonAction>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        foreach (var row in raw.Split(new[] { rowSep }, StringSplitOptions.None))
        {
            if (row.Trim().Length == 0) continue;

            var cols = row.Split('|');
            // Need at least a name plus one field (v1: UBound(cols) >= 1).
            if (cols.Length < 2) continue;

            var name = cols[0].Trim();
            if (name.Length == 0) continue;

            string script = string.Empty, input = string.Empty, output = string.Empty;

            if (cols[1].IndexOf('=') >= 0)
            {
                // Named format: key=value fields.
                for (int j = 1; j < cols.Length; j++)
                {
                    // Split on the FIRST '=' only — a value may contain '='.
                    var kv = cols[j].Split(new[] { '=' }, 2);
                    if (kv.Length < 2) continue;

                    var key = kv[0].Trim();
                    var val = kv[1].Trim();
                    switch (key)
                    {
                        case "script": script = val; break;
                        case "input": input = Accumulate(input, val); break;
                        case "output": output = Accumulate(output, val); break;
                        // entireRow / entreToEnd and any unknown key: designed
                        // out of v2 — intentionally ignored.
                    }
                }
            }
            else if (cols.Length >= 4)
            {
                // Legacy positional: name|script|input|output|[entireRow]|[entreToEnd].
                script = cols[1].Trim();
                input = cols[2].Trim();
                output = cols[3].Trim();
            }
            // else: a single non-named field with no positional payload — v1
            // produced an all-blank action here; carry the name forward the
            // same way (blank script/input/output).

            if (seen.Add(name))
                actions.Add(new RibbonAction(name, script, input, output));
        }

        return actions;
    }

    /// <summary>v1 joined repeated <c>input=</c> / <c>output=</c> fields with
    /// <c>"; "</c>; reproduce that so a round-trip through v1's own writer is
    /// preserved.</summary>
    private static string Accumulate(string current, string next)
        => current.Length == 0 ? next : current + "; " + next;

    /// <summary>True when <paramref name="legacy"/> carries any per-<em>sheet</em>
    /// PyExcel state (a script, a range field, an import/export/paste field, or
    /// at least one parseable action). The workbook-scoped
    /// <see cref="LegacyWorkbookState.Enabled"/> flag is deliberately ignored:
    /// the migration driver uses this to find the one sheet whose state to
    /// carry forward, and an enabled-but-otherwise-empty workbook has no sheet
    /// worth choosing on this basis.</summary>
    public static bool HasSheetContent(LegacyWorkbookState legacy)
    {
        if (legacy is null) throw new ArgumentNullException(nameof(legacy));
        return Optional(legacy.SelectedAction) is not null
            || Optional(legacy.SelectedScript) is not null
            || Optional(legacy.PyInput) is not null
            || Optional(legacy.PyOutput) is not null
            || Optional(legacy.ImportInput) is not null
            || Optional(legacy.ImportOutput) is not null
            || Optional(legacy.ExportInput) is not null
            || Optional(legacy.ExportOutput) is not null
            || Optional(legacy.PasteOutput) is not null
            || ParseActions(legacy.Actions).Count > 0;
    }

    /// <summary>Trim a v1 text field; an empty/blank result becomes
    /// <see langword="null"/> so the codec omits the element (v1's
    /// <c>GetSheetValue</c> returns <c>""</c> for an absent Name).</summary>
    private static string? Optional(string? value)
    {
        if (value is null) return null;
        var trimmed = value.Trim();
        return trimmed.Length == 0 ? null : trimmed;
    }

    /// <summary>v1 wrote <c>"1"</c> / <c>"0"</c> for the enabled flag; accept
    /// the textual booleans too for resilience. Anything else (including blank
    /// or absent) is <see langword="false"/>.</summary>
    private static bool ParseEnabled(string? value)
    {
        var v = value?.Trim();
        return string.Equals(v, "1", StringComparison.Ordinal)
            || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
    }
}

/// <summary>
/// The raw v1 defined-Name values captured for a single sheet, as the
/// Windows-only driver reads them off a live workbook. Every field is the
/// already-evaluated string (v1's <c>GetSheetValue</c> result) — the
/// <c>="…"</c> formula escaping is undone by Excel's evaluator before it
/// reaches here, exactly as in v1. All fields default to <see langword="null"/>
/// (an absent Name), so the driver only sets the ones the workbook actually
/// carries. See <see cref="LegacyStateConverter.LegacyNames"/> for the Name
/// each field comes from.
/// </summary>
public sealed record LegacyWorkbookState
{
    /// <summary><see cref="LegacyStateConverter.LegacyNames.Enabled"/>
    /// (workbook-scoped).</summary>
    public string? Enabled { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.SelectedAction"/>.</summary>
    public string? SelectedAction { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.Actions"/> —
    /// the serialized action list.</summary>
    public string? Actions { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.SelectedScript"/>.</summary>
    public string? SelectedScript { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.PyInput"/>.</summary>
    public string? PyInput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.PyOutput"/>.</summary>
    public string? PyOutput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.ImportInput"/>.</summary>
    public string? ImportInput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.ImportOutput"/>.</summary>
    public string? ImportOutput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.ExportInput"/>.</summary>
    public string? ExportInput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.ExportOutput"/>.</summary>
    public string? ExportOutput { get; init; }

    /// <summary><see cref="LegacyStateConverter.LegacyNames.PasteOutput"/>.</summary>
    public string? PasteOutput { get; init; }
}
