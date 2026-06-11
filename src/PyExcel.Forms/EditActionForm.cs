#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// The Add / Edit-action dialog (Phase 8 port of v1's <c>frmEditAction</c>,
/// reshaped onto the slimmer v2 <see cref="RibbonAction"/> model: a name,
/// a script, an input range, an output range, and optional keyword
/// arguments). The COM-bound ribbon shows it from <c>OnAddAction</c> /
/// <c>OnEditAction</c> and hands the result to <c>StateService.AddAction</c>.
///
/// <para>All field validation runs through the cross-platform
/// <see cref="EditActionValidator"/> and <see cref="KwargsText"/>, so the
/// dialog's decisions are tested on Linux CI; this class is only the
/// WinForms shell and so lives behind <c>#if NETFRAMEWORK</c>.</para>
///
/// <para>The form is shown modally with an owner (no off-screen
/// <c>(-20000,-20000)</c> hide hack from v1), validates on Save, and only
/// closes when the input is valid — invalid input is surfaced inline,
/// never pushed downstream.</para>
/// </summary>
public sealed class EditActionForm : Form
{
    private readonly TextBox _nameBox;
    private readonly ComboBox _scriptBox;
    private readonly TextBox _inputBox;
    private readonly TextBox _outputBox;
    private readonly TextBox _kwargsBox;
    private readonly Label _errorLabel;

    private readonly IReadOnlyList<string> _existingActionNames;
    private readonly string? _originalName;
    private readonly Func<string?>? _selectionProvider;

    /// <summary>The action the user built, valid only after the dialog
    /// returns <see cref="DialogResult.OK"/>.</summary>
    public RibbonAction? Result { get; private set; }

    /// <summary>
    /// Show the dialog modally and return the resulting action, or null if
    /// the user cancelled. The single entry point the ribbon calls.
    /// </summary>
    /// <param name="owner">The Excel main window to own the modal, so it
    /// can't be lost behind Excel or off-screen.</param>
    /// <param name="availableScripts">Scripts the user can pick from.</param>
    /// <param name="existingActionNames">Names already in the workbook,
    /// used to reject a duplicate name.</param>
    /// <param name="existing">The action being edited, or null to add a
    /// new one.</param>
    /// <param name="selectionProvider">Optional provider of the active
    /// Excel selection's address; when supplied, the input/output range
    /// fields get a "Pick…" button opening the range picker.</param>
    public static RibbonAction? Prompt(
        IWin32Window? owner,
        IReadOnlyList<string> availableScripts,
        IReadOnlyList<string> existingActionNames,
        RibbonAction? existing,
        Func<string?>? selectionProvider = null)
    {
        using var form = new EditActionForm(
            availableScripts, existingActionNames, existing, selectionProvider);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.Result : null;
    }

    private EditActionForm(
        IReadOnlyList<string> availableScripts,
        IReadOnlyList<string> existingActionNames,
        RibbonAction? existing,
        Func<string?>? selectionProvider)
    {
        _existingActionNames = existingActionNames ?? Array.Empty<string>();
        _originalName = existing?.Name;
        _selectionProvider = selectionProvider;

        Text = existing is null ? "Add Action" : "Edit Action";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(420, 360);
        Font = SystemFonts.MessageBoxFont;

        const int labelX = 12;
        const int fieldX = 110;
        const int fieldWidth = 296;
        var y = 14;

        AddLabel("Name:", labelX, y + 3);
        _nameBox = new TextBox { Left = fieldX, Top = y, Width = fieldWidth, TabIndex = 0 };
        Controls.Add(_nameBox);
        y += 32;

        AddLabel("Script:", labelX, y + 3);
        _scriptBox = new ComboBox
        {
            Left = fieldX,
            Top = y,
            Width = fieldWidth,
            DropDownStyle = ComboBoxStyle.DropDownList,
            TabIndex = 1,
        };
        foreach (var s in availableScripts ?? Array.Empty<string>())
            _scriptBox.Items.Add(s);
        Controls.Add(_scriptBox);
        y += 32;

        bool canPick = _selectionProvider is not null;
        int rangeWidth = canPick ? fieldWidth - 84 : fieldWidth;

        AddLabel("Input range:", labelX, y + 3);
        _inputBox = new TextBox { Left = fieldX, Top = y, Width = rangeWidth, TabIndex = 2 };
        Controls.Add(_inputBox);
        if (canPick) AddPickButton(_inputBox, fieldX + rangeWidth + 4, y - 1, tabIndex: 7);
        y += 32;

        AddLabel("Output range:", labelX, y + 3);
        _outputBox = new TextBox { Left = fieldX, Top = y, Width = rangeWidth, TabIndex = 3 };
        Controls.Add(_outputBox);
        if (canPick) AddPickButton(_outputBox, fieldX + rangeWidth + 4, y - 1, tabIndex: 8);
        y += 32;

        AddLabel("Keyword args:", labelX, y + 3);
        _kwargsBox = new TextBox
        {
            Left = fieldX,
            Top = y,
            Width = fieldWidth,
            Height = 96,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            AcceptsReturn = true,
            TabIndex = 4,
        };
        Controls.Add(_kwargsBox);
        AddLabel("one name=value per line", fieldX, y + 100, dim: true);
        y += 124;

        _errorLabel = new Label
        {
            Left = labelX,
            Top = y,
            Width = fieldX + fieldWidth - labelX,
            Height = 30,
            ForeColor = Color.Firebrick,
            Visible = false,
        };
        Controls.Add(_errorLabel);

        var saveButton = new Button
        {
            Text = "Save",
            DialogResult = DialogResult.None,
            Left = ClientSize.Width - 178,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 5,
        };
        saveButton.Click += OnSaveClick;
        Controls.Add(saveButton);

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 36,
            Width = 80,
            TabIndex = 6,
        };
        Controls.Add(cancelButton);

        AcceptButton = saveButton;
        CancelButton = cancelButton;

        if (existing is not null)
        {
            _nameBox.Text = existing.Name;
            _scriptBox.SelectedItem = existing.Script;
            _inputBox.Text = existing.Input;
            _outputBox.Text = existing.Output;
            _kwargsBox.Text = KwargsText.Format(existing.Kwargs);
        }
    }

    private void AddPickButton(TextBox target, int left, int top, int tabIndex)
    {
        var pick = new Button { Text = "Pick…", Left = left, Top = top, Width = 80, TabIndex = tabIndex };
        pick.Click += (_, _) =>
        {
            var picked = RangePickerForm.Prompt(this, target.Text, _selectionProvider);
            if (picked is not null) target.Text = picked;
        };
        Controls.Add(pick);
    }

    private void AddLabel(string text, int x, int top, bool dim = false)
    {
        Controls.Add(new Label
        {
            Text = text,
            Left = x,
            Top = top,
            AutoSize = true,
            ForeColor = dim ? SystemColors.GrayText : SystemColors.ControlText,
        });
    }

    private void OnSaveClick(object? sender, EventArgs e)
    {
        var kwargs = KwargsText.TryParse(_kwargsBox.Text, out var kwargsError);
        if (kwargs is null)
        {
            ShowError(kwargsError!);
            return;
        }

        var result = EditActionValidator.Validate(
            _nameBox.Text,
            _scriptBox.SelectedItem as string,
            _inputBox.Text,
            _outputBox.Text,
            kwargs,
            _existingActionNames,
            _originalName);

        if (!result.IsValid)
        {
            ShowError(result.ErrorMessage!);
            return;
        }

        Result = result.Action;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ShowError(string message)
    {
        _errorLabel.Text = message;
        _errorLabel.Visible = true;
    }
}
#endif
