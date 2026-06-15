#if NETFRAMEWORK
using System;
using System.Diagnostics;
using PyExcel.State;

namespace PyExcel.Addin;

// Declared inside the namespace deliberately — see the note in AppEventSink.cs.
// At file scope the bare name `Excel` binds to our own PyExcel.Excel namespace
// (CS0234) instead of this alias.
using Excel = Microsoft.Office.Interop.Excel;
// CustomXMLPart / CustomXMLParts (the types Workbook.CustomXMLParts actually
// returns) live in the Office object library, not the Excel interop — qualify
// them through Office.Core, which the project references via Office.dll.
using Core = Microsoft.Office.Core;

/// <summary>
/// The single, canonical persistence shell for a workbook's PyExcel profile:
/// reads and writes it as a <c>CustomXMLPart</c> embedded in the workbook, so
/// the profile travels inside the <c>.xlsx</c> and survives close/reopen, move,
/// rename, and cloud round-trips — with no sidecar file and no per-machine
/// app-data to keep in step. The workbook <em>is</em> the source of truth.
///
/// <para>The part carries the full <see cref="ProjectProfile"/> — the per-sheet
/// <see cref="WorkbookProfileData"/> plus environment <see cref="ProjectMetadata"/>
/// — serialised by the cross-platform <see cref="ProjectProfileCodec"/> and
/// located by <see cref="ProjectProfileCodec.XmlNamespace"/>. A profile written by
/// an earlier build as a state-only <see cref="WorkbookStateCodec"/> part (or a
/// flat single-state document nested in the project part) is still read and
/// migrated forward on the next save, so already-enabled workbooks keep their
/// configuration.</para>
///
/// <para>Both operations are best-effort: COM failures are logged, never thrown,
/// so neither a save hook nor an open hook can abort the user's action or
/// destabilise Excel.</para>
/// </summary>
internal static class WorkbookStatePersister
{
    // Any key works for the prior-metadata read below — it only affects the
    // WorkbookKey stamped on a migrated flat state, which we discard (we keep
    // only the metadata). Using a fixed sentinel keeps the intent clear.
    private const string MetadataProbeKey = "__pyexcel_prior__";

    /// <summary>Serialize <paramref name="data"/> — with fresh environment
    /// metadata, preserving the prior <c>created-utc</c> — and store it on
    /// <paramref name="workbook"/>, replacing any previously-saved PyExcel part.
    /// The part is added to the workbook's in-memory collection; it lands on disk
    /// when the workbook is saved (this method is called from the save hook).</summary>
    public static void Save(
        Excel.Workbook workbook, WorkbookProfileData data,
        string? projectDir, string? workbookName, string? workbookPath)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        if (data is null) throw new ArgumentNullException(nameof(data));
        try
        {
            // Preserve created-utc (and any non-recomputable field) from the
            // profile already on the workbook, if one is there.
            var prior = TryLoadProfile(workbook, MetadataProbeKey)?.Metadata;
            var meta = ProjectMetadataFactory.Build(projectDir, workbookName, workbookPath, prior);

            RemoveExisting(workbook);
            string xml = ProjectProfileCodec.SerializeToString(data, meta);
            workbook.CustomXMLParts.Add(xml);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"WorkbookStatePersister.Save failed: {ex}");
        }
    }

    /// <summary>Load the workbook profile from the embedded PyExcel part, or null
    /// when the workbook carries no readable PyExcel profile (one PyExcel has
    /// never touched) or the part is unreadable.</summary>
    public static WorkbookProfileData? TryLoad(Excel.Workbook workbook, string workbookKey)
        => TryLoadProfile(workbook, workbookKey)?.Profile;

    /// <summary>Load the full profile (per-sheet data + metadata) from the embedded
    /// part. Reads the current full-profile part first, then falls back to a
    /// state-only part written by an earlier build so those workbooks keep their
    /// state. <paramref name="workbookKey"/> only keys a migrated flat document.</summary>
    public static ProjectProfile? TryLoadProfile(Excel.Workbook workbook, string workbookKey)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        try
        {
            // Current format: the full profile, located by the project namespace.
            var xml = FirstPartXml(workbook, ProjectProfileCodec.XmlNamespace);
            if (xml is not null
                && ProjectProfileCodec.TryDeserialize(xml, workbookKey, out var data, out var meta)
                && data is not null)
            {
                return new ProjectProfile(data, meta ?? new ProjectMetadata());
            }

            // Legacy format: a state-only part from a build before the profile
            // existed — migrate its flat single state forward into the per-sheet
            // model so those workbooks keep their configuration.
            var legacy = FirstPartXml(workbook, WorkbookStateCodec.XmlNamespace);
            if (legacy is not null
                && WorkbookStateCodec.TryDeserialize(legacy, workbookKey, out var legacyState)
                && legacyState is not null)
            {
                return new ProjectProfile(WorkbookProfileData.FromState(legacyState), new ProjectMetadata());
            }

            return null;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"WorkbookStatePersister.TryLoadProfile failed: {ex}");
            return null;
        }
    }

    /// <summary>The XML of the first custom part in namespace <paramref name="ns"/>,
    /// or <see langword="null"/> if the workbook has none there.</summary>
    private static string? FirstPartXml(Excel.Workbook workbook, string ns)
    {
        Core.CustomXMLParts parts = workbook.CustomXMLParts.SelectByNamespace(ns);
        if (parts is null || parts.Count == 0) return null;
        // The Excel object model is 1-based. Save removes before adding, so there
        // is only ever one PyExcel part per namespace — take the first.
        return parts[1].XML;
    }

    /// <summary>Delete every PyExcel part — both the current full-profile
    /// namespace and the legacy state-only namespace — so a re-save never
    /// accumulates duplicates or leaves a stale state-only part behind.</summary>
    private static void RemoveExisting(Excel.Workbook workbook)
    {
        foreach (var ns in new[] { ProjectProfileCodec.XmlNamespace, WorkbookStateCodec.XmlNamespace })
        {
            Core.CustomXMLParts existing = workbook.CustomXMLParts.SelectByNamespace(ns);
            if (existing is null) continue;
            // Walk back-to-front: the collection re-indexes as parts are
            // removed, so a forward walk would skip elements.
            for (int i = existing.Count; i >= 1; i--)
                existing[i].Delete();
        }
    }
}
#endif
