#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// Helper for invoking the injected <b>native</b> Excel range picker
/// (Application.InputBox with Type:=8) from inside a modal dialog.
///
/// <para>Excel's range selector needs the user to click on the sheet, so any
/// PyExcel dialog covering the grid must step out of the way for the duration
/// of the pick. This hides the supplied dialog(s), runs the picker, then
/// restores and re-focuses them — the fix for the old "Pick" button, which
/// could only echo back whatever was selected <em>before</em> the dialog
/// opened.</para>
/// </summary>
internal static class RangePick
{
    /// <summary>
    /// Hide <paramref name="dialogs"/> so Excel is interactive, invoke
    /// <paramref name="picker"/> with <paramref name="initial"/>, then restore
    /// them. Returns the picked address, or null when there is no picker or the
    /// user cancelled. The first dialog passed is re-activated afterwards, so
    /// pass the top-most form first.
    /// </summary>
    public static string? OnSheet(
        Func<string?, string?>? picker, string? initial, params Form?[] dialogs)
    {
        if (picker is null) return null;

        var hidden = new List<Form>();
        foreach (var f in dialogs)
        {
            if (f is not null && !f.IsDisposed && f.Visible)
            {
                f.Visible = false;
                hidden.Add(f);
            }
        }

        try
        {
            return picker(initial);
        }
        finally
        {
            foreach (var f in hidden)
            {
                if (!f.IsDisposed) f.Visible = true;
            }
            var first = dialogs.Length > 0 ? dialogs[0] : null;
            if (first is not null && !first.IsDisposed) first.Activate();
        }
    }
}
#endif
