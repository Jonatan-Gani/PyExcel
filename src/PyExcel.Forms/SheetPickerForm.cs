#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace PyExcel.Forms;

/// <summary>
/// The sheet-picker dialog (Phase 8 port of v1's <c>SheetPickerForm</c>):
/// a drop-down of the workbook's sheet names with OK / Cancel. The
/// Excel-import flow shows it when the user gave an Excel path without a
/// pinned <c>!Sheet</c> and the workbook has more than one sheet.
///
/// <para>Selection validation runs through the cross-platform
/// <see cref="SheetPickerValidator"/>, so this class is only the WinForms
/// shell and lives behind <c>#if NETFRAMEWORK</c>. Shown modally with an
/// owner (no off-screen hide hack); OK closes only when a sheet is
/// chosen.</para>
/// </summary>
public sealed class SheetPickerForm : Form
{
    private readonly ComboBox _sheetBox;
    private readonly Label _errorLabel;
    private readonly IReadOnlyList<string> _availableSheets;

    /// <summary>The chosen sheet (canonical casing), valid only after the
    /// dialog returns <see cref="DialogResult.OK"/>.</summary>
    public string? SelectedSheet { get; private set; }

    /// <summary>
    /// Show the picker modally and return the chosen sheet, or null if the
    /// user cancelled. The single entry point the import flow calls.
    /// </summary>
    /// <param name="owner">Excel's main window, so the modal can't be lost
    /// behind Excel or off-screen.</param>
    /// <param name="availableSheets">The workbook's sheet names.</param>
    /// <param name="preselected">The sheet to pre-select, or null for the
    /// first.</param>
    public static string? Prompt(
        IWin32Window? owner,
        IReadOnlyList<string> availableSheets,
        string? preselected)
    {
        using var form = new SheetPickerForm(availableSheets, preselected);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.SelectedSheet : null;
    }

    private SheetPickerForm(IReadOnlyList<string> availableSheets, string? preselected)
    {
        _availableSheets = availableSheets ?? Array.Empty<string>();

        Text = "Pick sheet";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(320, 132);
        Font = SystemFonts.MessageBoxFont;

        Controls.Add(new Label
        {
            Text = "Sheet:",
            Left = 12,
            Top = 17,
            AutoSize = true,
        });

        _sheetBox = new ComboBox
        {
            Left = 64,
            Top = 14,
            Width = 244,
            DropDownStyle = ComboBoxStyle.DropDownList,
            TabIndex = 0,
        };
        foreach (var sheet in _availableSheets)
            _sheetBox.Items.Add(sheet);
        if (preselected is not null && _sheetBox.Items.Contains(preselected))
            _sheetBox.SelectedItem = preselected;
        else if (_sheetBox.Items.Count > 0)
            _sheetBox.SelectedIndex = 0;
        Controls.Add(_sheetBox);

        _errorLabel = new Label
        {
            Left = 12,
            Top = 48,
            Width = 296,
            Height = 30,
            ForeColor = Color.Firebrick,
            Visible = false,
        };
        Controls.Add(_errorLabel);

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.None,
            Left = ClientSize.Width - 178,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 1,
        };
        okButton.Click += OnOkClick;
        Controls.Add(okButton);

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 2,
        };
        Controls.Add(cancelButton);

        AcceptButton = okButton;
        CancelButton = cancelButton;
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var result = SheetPickerValidator.Validate(
            _sheetBox.SelectedItem as string, _availableSheets);
        if (!result.IsValid)
        {
            _errorLabel.Text = result.ErrorMessage;
            _errorLabel.Visible = true;
            return;
        }

        SelectedSheet = result.SelectedSheet;
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
