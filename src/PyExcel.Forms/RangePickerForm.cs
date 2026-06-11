#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// The range-picker dialog (Phase 8). A small editor for a single range
/// reference with an optional "Use current selection" button that pulls
/// the active Excel selection's address through an injected provider — so
/// this WinForms shell takes no COM dependency (the ribbon supplies the
/// provider, reading <c>Application.Selection</c> late-bound).
///
/// <para>Validation runs through the cross-platform
/// <see cref="RangeAddressValidator"/>; OK closes only on a valid single
/// range.</para>
/// </summary>
public sealed class RangePickerForm : Form
{
    private readonly TextBox _addressBox;
    private readonly Label _errorLabel;

    public string? SelectedAddress { get; private set; }

    /// <summary>Show the picker and return the chosen range, or null on
    /// cancel.</summary>
    /// <param name="owner">Excel's main window.</param>
    /// <param name="initial">The address to pre-fill.</param>
    /// <param name="selectionProvider">Returns the active selection's
    /// address, or null if unavailable; when null the "Use current
    /// selection" button is hidden.</param>
    public static string? Prompt(
        IWin32Window? owner, string? initial, Func<string?>? selectionProvider)
    {
        using var form = new RangePickerForm(initial, selectionProvider);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.SelectedAddress : null;
    }

    private RangePickerForm(string? initial, Func<string?>? selectionProvider)
    {
        Text = "Pick range";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(360, 128);

        Controls.Add(new Label { Text = "Range:", Left = 12, Top = 17, AutoSize = true });

        bool canPick = selectionProvider is not null;
        _addressBox = new TextBox
        {
            Left = 64,
            Top = 14,
            Width = canPick ? 200 : 284,
            Text = initial ?? string.Empty,
            TabIndex = 0,
        };
        Controls.Add(_addressBox);

        if (canPick)
        {
            var useSelection = new Button
            {
                Text = "Use selection",
                Left = 270,
                Top = 13,
                Width = 78,
                TabIndex = 1,
            };
            useSelection.Click += (_, _) =>
            {
                var addr = selectionProvider!();
                if (!string.IsNullOrEmpty(addr)) _addressBox.Text = addr;
            };
            Controls.Add(useSelection);
        }

        _errorLabel = new Label
        {
            Left = 12,
            Top = 46,
            Width = 336,
            Height = 30,
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
            TabIndex = 2,
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
            TabIndex = 3,
        };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var result = RangeAddressValidator.Validate(_addressBox.Text);
        if (!result.IsValid)
        {
            _errorLabel.Text = result.ErrorMessage;
            _errorLabel.Visible = true;
            return;
        }

        SelectedAddress = result.Address;
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
