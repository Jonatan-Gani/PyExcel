#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// A minimal single-line text prompt — WinForms has no built-in InputBox.
/// Used to ask for a new script's name (Note 2). Returns the trimmed value
/// on OK, or null on Cancel.
/// </summary>
internal sealed class TextPromptForm : ScaledForm
{
    private readonly TextBox _box;

    public string? Value { get; private set; }

    public static string? Prompt(
        IWin32Window? owner, string title, string label, string? initial = null)
    {
        using var form = new TextPromptForm(title, label, initial);
        var r = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return r == DialogResult.OK ? form.Value : null;
    }

    private TextPromptForm(string title, string label, string? initial)
    {
        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(360, 96);

        Controls.Add(new Label { Text = label, Left = 12, Top = 14, AutoSize = true });
        _box = new TextBox
        {
            Left = 12,
            Top = 36,
            Width = 336,
            Text = initial ?? string.Empty,
            TabIndex = 0,
        };
        Controls.Add(_box);

        var ok = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Left = ClientSize.Width - 178,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 1,
        };
        ok.Click += (_, _) => Value = _box.Text.Trim();
        Controls.Add(ok);

        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 2,
        };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
    }
}
#endif
