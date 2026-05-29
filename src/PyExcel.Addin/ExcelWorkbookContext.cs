using System;
using ExcelDna.Integration;
using PyExcel.State;

namespace PyExcel.Addin;

/// <summary>
/// Concrete <see cref="IWorkbookContext"/> backed by
/// <c>Application.ActiveWorkbook</c> over the Excel COM interop. Wired
/// into <see cref="PyExcelServices.WorkbookContext"/> by
/// <see cref="AddIn.AutoOpen"/>; tests use a hand-written fake instead
/// of this class so the cross-platform CI slice doesn't need Excel.
///
/// <para>Key strategy:</para>
/// <list type="bullet">
///   <item>Saved workbooks → <c>Workbook.FullName</c> (the full path
///     including the filename). Stable across closes/reopens because
///     the path is the on-disk identity.</item>
///   <item>Unsaved workbooks (no on-disk path) → a synthetic
///     <c>"unsaved:{SessionGuid}:{Workbook.Name}"</c> key. Excel
///     gives every new workbook a unique <c>Name</c> within a session
///     (<c>Book1</c>, <c>Book2</c>, …), and the session GUID is
///     allocated once per add-in load — so two unsaved workbooks in
///     the same session don't collide, and a "Save As" promotes the
///     workbook to a path-based key on the next access (the previous
///     unsaved key is then orphaned in the <see cref="StateService"/>
///     until <c>WorkbookBeforeClose</c> calls
///     <see cref="StateService.Forget"/>).</item>
/// </list>
///
/// <para>Returns <see langword="null"/> when Excel has no active
/// workbook — typically during add-in load before any workbook is
/// open. The ribbon renders as "all disabled" in that state, which is
/// what we want.</para>
/// </summary>
public sealed class ExcelWorkbookContext : IWorkbookContext
{
    // One GUID per add-in load. Distinguishes "Book1 from this Excel
    // session" from "Book1 from a previous session whose state lingered
    // somewhere it shouldn't" — defensive, since session state should
    // not outlive the add-in instance anyway.
    private static readonly string SessionGuid = Guid.NewGuid().ToString("N");

    public string? CurrentWorkbookKey
    {
        get
        {
            try
            {
                // dynamic instead of the typed COM PIA: keeps PyExcel.Addin
                // free of an Office interop assembly reference, matching
                // the Excel-DNA convention.
                dynamic app = ExcelDnaUtil.Application;
                dynamic wb = app.ActiveWorkbook;
                if (wb is null) return null;

                string name = (string)wb.Name;
                string path = (string)wb.Path;

                // Excel returns an empty Path for new-but-unsaved workbooks.
                if (string.IsNullOrEmpty(path))
                    return $"unsaved:{SessionGuid}:{name}";

                return (string)wb.FullName;
            }
            catch
            {
                // COM exceptions during shutdown / between workbook events
                // can leave ActiveWorkbook in a transient state. The
                // ribbon's getters all tolerate a null key (they render as
                // "no workbook"), so a swallow-and-return-null here is the
                // right answer rather than crashing the ribbon callback.
                return null;
            }
        }
    }
}
