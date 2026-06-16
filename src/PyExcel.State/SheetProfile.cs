using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// The per-<em>sheet</em> slice of a PyExcel project: everything the user
/// configures while working on one worksheet — the selected script, the Run
/// input/output bindings, the saved actions, and the Import/Export/Paste field
/// bindings. A <see cref="WorkbookState"/> projects exactly one of these (the
/// active sheet's) onto its flat fields, so the ribbon and the data services keep
/// reading <c>state.SelectedScript</c> etc. without knowing sheets exist.
///
/// <para>Workbook-scoped facts — whether the workbook is enabled, its project
/// directory, the shared userScripts list — are <em>not</em> here; they live on
/// <see cref="WorkbookProfileData"/> / <see cref="WorkbookState"/>.</para>
/// </summary>
public sealed record SheetProfile
{
    public string? SelectedScript { get; init; }
    public string? PyInput { get; init; }
    public string? PyOutput { get; init; }
    public IReadOnlyList<RibbonAction> Actions { get; init; } = Array.Empty<RibbonAction>();
    public string? SelectedActionName { get; init; }
    public string? ImportInput { get; init; }
    public string? ImportOutput { get; init; }
    public string? ExportInput { get; init; }
    public string? ExportOutput { get; init; }

    /// <summary>Export default: the destination folder (blank → next to the
    /// workbook).</summary>
    public string? ExportFolder { get; init; }

    /// <summary>Export default: the base file name (no extension or stamp).</summary>
    public string? ExportBaseName { get; init; }

    /// <summary>Export default: the file-type token (<c>csv</c> / <c>tsv</c>).</summary>
    public string? ExportFormat { get; init; }

    /// <summary>Export default: the unique-name stamp token (<c>none</c> /
    /// <c>datetime</c> / <c>date</c> / <c>compact</c>).</summary>
    public string? ExportTimestamp { get; init; }

    public string? PasteOutput { get; init; }

    /// <summary>The all-empty profile — an unconfigured sheet.</summary>
    public static readonly SheetProfile Empty = new();

    /// <summary>The selected action resolved against <see cref="Actions"/>, or
    /// <see langword="null"/> if nothing is selected (or it doesn't resolve).</summary>
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

    /// <summary>True when the sheet carries any user configuration worth saving.
    /// Used to skip persisting (and inheriting) untouched sheets so the on-disk
    /// profile stays lean.</summary>
    public bool IsConfigured =>
        Actions.Count > 0
        || SelectedScript is not null
        || PyInput is not null
        || PyOutput is not null
        || SelectedActionName is not null
        || ImportInput is not null
        || ImportOutput is not null
        || ExportInput is not null
        || ExportOutput is not null
        || ExportFolder is not null
        || ExportBaseName is not null
        || ExportFormat is not null
        || ExportTimestamp is not null
        || PasteOutput is not null;
}
