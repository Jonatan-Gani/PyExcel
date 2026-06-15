#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// Read-only, resizable viewer for the last captured kernel error, used by
/// the ribbon's "Show Last Error" button.
///
/// <para>It replaces the previous reliance on Excel-DNA's
/// <c>LogDisplay.Show()</c>, which is a no-op once the log window is already
/// open behind Excel — and never raises it to the foreground — so the button
/// appeared to "do nothing". This dialog is shown modal and owned by Excel's
/// window, so it always comes to the front and the user can read (and copy)
/// the traceback.</para>
/// </summary>
public sealed class ErrorDisplayForm : ScaledForm
{
    /// <summary>Show the error <paramref name="body"/> modally, owned by
    /// <paramref name="owner"/> (Excel's main window) so it is always
    /// front-most.</summary>
    public static void Open(IWin32Window? owner, string title, string body)
    {
        using var form = new ErrorDisplayForm(title, body);
        if (owner is null) form.ShowDialog(); else form.ShowDialog(owner);
    }

    private ErrorDisplayForm(string title, string body)
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.Sizable;
        StartPosition = FormStartPosition.CenterParent;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(560, 360);
        MinimumSize = new Size(360, 220);

        var box = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            HideSelection = true,
            Left = 12,
            Top = 12,
            Width = ClientSize.Width - 24,
            Height = ClientSize.Height - 56,
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom
                   | AnchorStyles.Left | AnchorStyles.Right,
            // A monospaced font keeps the Python traceback's indentation
            // readable, the same way it renders in a terminal.
            Font = new Font(FontFamily.GenericMonospace, 9f),
            Text = body,
        };
        Controls.Add(box);

        var copy = new Button
        {
            Text = "Copy",
            Width = 80,
            Height = 28,
            Left = ClientSize.Width - 180,
            Top = ClientSize.Height - 40,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
        };
        // Clipboard needs an STA thread; ShowDialog runs on Excel's STA main
        // thread, so this is safe. Swallow the rare clipboard-busy failure.
        copy.Click += (_, _) => { try { Clipboard.SetText(body); } catch { /* best-effort */ } };
        Controls.Add(copy);

        var close = new Button
        {
            Text = "Close",
            Width = 80,
            Height = 28,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 40,
            DialogResult = DialogResult.OK,
            Anchor = AnchorStyles.Bottom | AnchorStyles.Right,
        };
        Controls.Add(close);

        AcceptButton = close;
        CancelButton = close;
    }
}
#endif
