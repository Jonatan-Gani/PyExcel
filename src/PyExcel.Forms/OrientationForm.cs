#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.Excel;

namespace PyExcel.Forms;

/// <summary>
/// The list-direction dialog (Phase 8 port of v1's <c>frmOrientation</c>).
/// Shown when a 1-D result spills into a single-cell target, where the
/// orientation is ambiguous — the user picks Across (a row) or Down (a
/// column). Returns null if they close without choosing.
/// </summary>
public sealed class OrientationForm : ScaledForm
{
    public ListOrientation? Choice { get; private set; }

    /// <summary>Show the dialog and return the chosen orientation, or null
    /// if cancelled.</summary>
    public static ListOrientation? Prompt(IWin32Window? owner)
    {
        using var form = new OrientationForm();
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.Choice : null;
    }

    private OrientationForm()
    {
        Text = "List direction";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(320, 96);

        Controls.Add(new Label
        {
            Text = "Spill the list across a row or down a column?",
            Left = 12,
            Top = 14,
            AutoSize = true,
        });

        var horizontal = new Button
        {
            Text = "Across (row)",
            Left = 12,
            Top = 48,
            Width = 140,
            TabIndex = 0,
        };
        horizontal.Click += (_, _) => Choose(ListOrientation.Horizontal);
        Controls.Add(horizontal);

        var vertical = new Button
        {
            Text = "Down (column)",
            Left = 168,
            Top = 48,
            Width = 140,
            TabIndex = 1,
        };
        vertical.Click += (_, _) => Choose(ListOrientation.Vertical);
        Controls.Add(vertical);

        AcceptButton = horizontal;
    }

    private void Choose(ListOrientation orientation)
    {
        Choice = orientation;
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
