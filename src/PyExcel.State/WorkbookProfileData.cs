using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// The full, persisted shape of a workbook's PyExcel project: the
/// workbook-scoped flags (<see cref="Enabled"/>, <see cref="ProjectDir"/>) plus
/// the per-sheet <see cref="SheetProfile"/> map. This is the transport/storage
/// form — <see cref="WorkbookProfileCodec"/> round-trips it into the workbook's
/// <c>CustomXMLPart</c>, and <see cref="StateService"/> exposes it via
/// <see cref="StateService.GetProfile"/> / loads it via
/// <see cref="StateService.LoadProfile"/>. The ribbon and the data services never
/// see it; they read the single-sheet <see cref="WorkbookState"/> projection.
/// </summary>
public sealed record WorkbookProfileData
{
    public bool Enabled { get; init; }
    public string? ProjectDir { get; init; }

    /// <summary>Stable project identity — a GUID stamped into the workbook on Enable
    /// (and lazily on the next save for projects enabled before this existed). Unlike
    /// the file-path key, it survives a move / rename / cloud round-trip, and lets a
    /// <em>copy</em> of an enabled workbook be told apart from the moved original
    /// (see <see cref="WorkbookIdentityReconciler"/>). Null on a workbook PyExcel has
    /// never enabled.</summary>
    public string? ProjectId { get; init; }

    /// <summary>The workbook's full path the last time its identity was committed
    /// (Enable, or a reconcile after a move). Compared against the workbook's current
    /// path on open to detect a move (origin gone) versus a copy (origin still on
    /// disk). Null when unstamped.</summary>
    public string? OriginPath { get; init; }

    /// <summary>Per-sheet configuration, keyed by worksheet name. The
    /// <see cref="DefaultSheetKey"/> entry is a workbook-level default a sheet
    /// inherits when it has no entry of its own — it is how a pre-per-sheet
    /// profile migrates forward (its single flat configuration becomes the
    /// default) and how a v1 workbook's one carried-over sheet lands.</summary>
    public IReadOnlyDictionary<string, SheetProfile> Sheets { get; init; }
        = new Dictionary<string, SheetProfile>(StringComparer.Ordinal);

    public static readonly WorkbookProfileData Empty = new();

    /// <summary>The default-bucket key (an empty string — never a real worksheet
    /// name): a workbook-level profile that sheets without their own entry
    /// inherit.</summary>
    public const string DefaultSheetKey = "";

    /// <summary>True when this carries anything worth persisting / restoring — it's
    /// enabled, has a project dir, or any sheet is configured. The signal the
    /// event sink uses to decide "is this workbook a PyExcel project?".</summary>
    public bool IsMeaningful
    {
        get
        {
            if (Enabled || !string.IsNullOrEmpty(ProjectDir)) return true;
            foreach (var kv in Sheets)
                if (kv.Value.IsConfigured) return true;
            return false;
        }
    }

    /// <summary>Build a profile from a single flat <see cref="WorkbookState"/> —
    /// its workbook-scoped fields stay workbook-scoped and its sheet-scoped fields
    /// become the <see cref="DefaultSheetKey"/> bucket. Used to carry a v1
    /// (single-state) or a pre-per-sheet (flat) profile forward into the per-sheet
    /// model without losing the user's configuration.</summary>
    public static WorkbookProfileData FromState(WorkbookState s)
    {
        if (s is null) throw new ArgumentNullException(nameof(s));
        var sheet = new SheetProfile
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
            ExportFolder = s.ExportFolder,
            ExportBaseName = s.ExportBaseName,
            ExportFormat = s.ExportFormat,
            ExportTimestamp = s.ExportTimestamp,
            PasteOutput = s.PasteOutput,
        };
        var sheets = new Dictionary<string, SheetProfile>(StringComparer.Ordinal);
        // Only seed the default bucket when the flat state actually had sheet
        // configuration — an enabled-but-unconfigured workbook gets no default.
        if (sheet.IsConfigured) sheets[DefaultSheetKey] = sheet;
        return new WorkbookProfileData
        {
            Enabled = s.Enabled,
            ProjectDir = s.ProjectDir,
            ProjectId = s.ProjectId,
            OriginPath = s.OriginPath,
            Sheets = sheets,
        };
    }
}
