#if NETFRAMEWORK
using System;
using System.Diagnostics;
using PyExcel.State;

namespace PyExcel.Addin;

// Declared inside the namespace deliberately — see the note in AppEventSink.cs.
// At file scope the bare name `Excel` binds to our own PyExcel.Excel namespace
// (CS0234) instead of this alias.
using Excel = Microsoft.Office.Interop.Excel;

/// <summary>
/// Reads and writes PyExcel's per-workbook state as a <c>CustomXMLPart</c>
/// attached to the workbook, so the state travels inside the file and
/// survives close/reopen. The XML round-trip itself is owned by the
/// cross-platform <see cref="WorkbookStateCodec"/> (unit-tested on Linux);
/// this class is the thin, Windows-only COM shell that locates / replaces /
/// reads the part on a live <see cref="Excel.Workbook"/>.
///
/// <para>The part is identified by <see cref="WorkbookStateCodec.XmlNamespace"/>
/// so PyExcel's part can be found among any other custom parts the workbook
/// carries (Office stores its own parts there too). Saving first deletes
/// every existing part in that namespace, then adds the current one — so a
/// re-save never accumulates duplicates.</para>
///
/// <para>Both operations are best-effort: COM failures are logged, never
/// thrown, so neither a save hook nor an open hook can abort the user's
/// action or destabilise Excel.</para>
/// </summary>
internal static class WorkbookStatePersister
{
    /// <summary>Serialize <paramref name="state"/> and store it on
    /// <paramref name="workbook"/>, replacing any previously-saved PyExcel
    /// part.</summary>
    public static void Save(Excel.Workbook workbook, WorkbookState state)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        if (state is null) throw new ArgumentNullException(nameof(state));
        try
        {
            RemoveExisting(workbook);
            string xml = WorkbookStateCodec.SerializeToString(state);
            workbook.CustomXMLParts.Add(xml);
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"WorkbookStatePersister.Save failed: {ex}");
        }
    }

    /// <summary>Find PyExcel's part on <paramref name="workbook"/> and
    /// deserialize it into a state keyed by <paramref name="workbookKey"/>.
    /// Returns <see langword="null"/> when there's no PyExcel part (a
    /// workbook PyExcel has never touched) or the part is unreadable.</summary>
    public static WorkbookState? TryLoad(Excel.Workbook workbook, string workbookKey)
    {
        if (workbook is null) throw new ArgumentNullException(nameof(workbook));
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        try
        {
            Excel.CustomXMLParts parts =
                workbook.CustomXMLParts.SelectByNamespace(WorkbookStateCodec.XmlNamespace);
            if (parts is null || parts.Count == 0) return null;

            // The Excel object model is 1-based. There should only ever be
            // one PyExcel part (Save deletes before adding); take the first.
            Excel.CustomXMLPart part = parts[1];
            return WorkbookStateCodec.TryDeserialize(part.XML, workbookKey, out var state)
                ? state
                : null;
        }
        catch (Exception ex)
        {
            Trace.WriteLine($"WorkbookStatePersister.TryLoad failed: {ex}");
            return null;
        }
    }

    private static void RemoveExisting(Excel.Workbook workbook)
    {
        Excel.CustomXMLParts existing =
            workbook.CustomXMLParts.SelectByNamespace(WorkbookStateCodec.XmlNamespace);
        if (existing is null) return;
        // Walk back-to-front: the collection re-indexes as parts are
        // removed, so a forward walk would skip elements.
        for (int i = existing.Count; i >= 1; i--)
            existing[i].Delete();
    }
}
#endif
