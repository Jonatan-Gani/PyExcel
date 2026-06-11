#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// Small editor for one range binding — a range reference plus an optional
/// name — opened by the Add / Edit buttons of <see cref="RangeListEditor"/>
/// (Note 4). The range is chosen with Excel's NATIVE picker (the "Pick…"
/// button) and validated through <see cref="RangeAddressValidator"/>; the
/// name is optional (blank means an anonymous, positional range).
/// </summary>
internal sealed class RangeNameForm : Form
{
    private readonly TextBox _rangeBox;
    private readonly TextBox _nameBox;
    private readonly Label _errorLabel;
    private readonly Func<string?, string?>? _rangePicker;

    public RangeBinding? Result { get; private set; }

    public static RangeBinding? Prompt(
        IWin32Window? owner, RangeBinding? initial, Func<string?, string?>? rangePicker)
    {
        using var form = new RangeNameForm(initial, rangePicker);
        var r = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return r == DialogResult.OK ? form.Result : null;
    }

    private RangeNameForm(RangeBinding? initial, Func<string?, string?>? rangePicker)
    {
        _rangePicker = rangePicker;

        Text = "Range";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(380, 152);

        Controls.Add(new Label { Text = "Range:", Left = 12, Top = 17, AutoSize = true });
        bool canPick = _rangePicker is not null;
        _rangeBox = new TextBox
        {
            Left = 80,
            Top = 14,
            Width = canPick ? 196 : 288,
            Text = initial?.RangeText ?? string.Empty,
            TabIndex = 0,
        };
        Controls.Add(_rangeBox);
        if (canPick)
        {
            var pick = new Button { Text = "Pick…", Left = 282, Top = 13, Width = 86, TabIndex = 1 };
            pick.Click += (_, _) => PickOnSheet();
            Controls.Add(pick);
        }

        Controls.Add(new Label { Text = "Name:", Left = 12, Top = 49, AutoSize = true });
        _nameBox = new TextBox
        {
            Left = 80,
            Top = 46,
            Width = 288,
            Text = initial?.Name ?? string.Empty,
            TabIndex = 2,
        };
        Controls.Add(_nameBox);
        Controls.Add(new Label
        {
            Text = "optional — leave blank for a positional range",
            Left = 80,
            Top = 70,
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
        });

        _errorLabel = new Label
        {
            Left = 12,
            Top = 94,
            Width = 356,
            Height = 28,
            ForeColor = Color.Firebrick,
            Visible = false,
        };
        Controls.Add(_errorLabel);

        var ok = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.None,
            Left = ClientSize.Width - 178,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 3,
        };
        ok.Click += OnOkClick;
        Controls.Add(ok);

        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 4,
        };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
    }

    private void PickOnSheet()
    {
        // Hide this dialog AND its owner (the EditAction form) so Excel is
        // fully interactive for the native range selector, then restore them.
        var picked = RangePick.OnSheet(_rangePicker, _rangeBox.Text, this, Owner);
        if (picked is not null) _rangeBox.Text = picked;
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var result = RangeAddressValidator.Validate(_rangeBox.Text);
        if (!result.IsValid)
        {
            _errorLabel.Text = result.ErrorMessage;
            _errorLabel.Visible = true;
            return;
        }

        var name = _nameBox.Text.Trim();
        Result = new RangeBinding(name.Length == 0 ? null : name, result.Address!);
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
