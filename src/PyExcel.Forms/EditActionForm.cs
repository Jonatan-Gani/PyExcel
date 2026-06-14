#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// The Add / Edit-action dialog (Phase 8 port of v1's <c>frmEditAction</c>,
/// reshaped onto the slimmer v2 <see cref="RibbonAction"/> model: a name, a
/// script, input ranges, output ranges, and optional keyword arguments).
///
/// <para>Notes 2 &amp; 4: the input/output ranges are edited as lists
/// (<see cref="RangeListEditor"/> — add/edit/remove/reorder, each row picked
/// with Excel's native range selector and optionally named), and the Script
/// field has a "New…" button (<see cref="ScriptScaffold"/>) so a fresh
/// workbook with no scripts isn't a dead end.</para>
///
/// <para>Field validation runs through the cross-platform
/// <see cref="EditActionValidator"/> and <see cref="KwargsText"/>; the list
/// editors serialise back to the same <c>{name}=range; …</c> string the
/// validator already expects, so the decisions stay unit-tested on Linux CI.
/// This class is only the WinForms shell, behind <c>#if NETFRAMEWORK</c>.</para>
/// </summary>
public sealed class EditActionForm : ScaledForm
{
    private readonly TextBox _nameBox;
    private readonly ComboBox _scriptBox;
    private readonly RangeListEditor _inputEditor;
    private readonly RangeListEditor _outputEditor;
    private readonly TextBox _kwargsBox;
    private readonly Label _errorLabel;

    private readonly IReadOnlyList<string> _existingActionNames;
    private readonly string? _originalName;
    private readonly string? _userScriptsDirectory;

    /// <summary>The action the user built, valid only after the dialog
    /// returns <see cref="DialogResult.OK"/>.</summary>
    public RibbonAction? Result { get; private set; }

    /// <summary>
    /// Show the dialog modally and return the resulting action, or null if the
    /// user cancelled.
    /// </summary>
    /// <param name="owner">Excel's main window, so the modal can't be lost.</param>
    /// <param name="availableScripts">Scripts the user can pick from.</param>
    /// <param name="existingActionNames">Names already in the workbook, used
    /// to reject a duplicate.</param>
    /// <param name="existing">The action being edited, or null to add.</param>
    /// <param name="rangePicker">Native range picker (initial → picked
    /// address). When supplied, the range rows get a "Pick…" button.</param>
    /// <param name="userScriptsDirectory">The workbook's <c>userScripts</c>
    /// folder. When supplied, the "New…" script button is enabled.</param>
    public static RibbonAction? Prompt(
        IWin32Window? owner,
        IReadOnlyList<string> availableScripts,
        IReadOnlyList<string> existingActionNames,
        RibbonAction? existing,
        Func<string?, string?>? rangePicker = null,
        string? userScriptsDirectory = null)
    {
        using var form = new EditActionForm(
            availableScripts, existingActionNames, existing, rangePicker, userScriptsDirectory);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.Result : null;
    }

    private EditActionForm(
        IReadOnlyList<string> availableScripts,
        IReadOnlyList<string> existingActionNames,
        RibbonAction? existing,
        Func<string?, string?>? rangePicker,
        string? userScriptsDirectory)
    {
        _existingActionNames = existingActionNames ?? Array.Empty<string>();
        _originalName = existing?.Name;
        _userScriptsDirectory = userScriptsDirectory;

        Text = existing is null ? "Add Action" : "Edit Action";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(440, 500);
        Font = SystemFonts.MessageBoxFont;

        const int labelX = 12;
        const int fieldX = 110;
        const int fieldWidth = 316;
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
            Width = fieldWidth - 84,
            DropDownStyle = ComboBoxStyle.DropDownList,
            TabIndex = 1,
        };
        foreach (var s in availableScripts ?? Array.Empty<string>())
            _scriptBox.Items.Add(s);
        Controls.Add(_scriptBox);
        var newScript = new Button
        {
            Text = "New…",
            Left = fieldX + fieldWidth - 80,
            Top = y - 1,
            Width = 80,
            TabIndex = 2,
            Enabled = _userScriptsDirectory is not null,
        };
        newScript.Click += (_, _) => OnNewScript();
        Controls.Add(newScript);
        y += 34;

        AddLabel("Input ranges:", labelX, y + 3);
        _inputEditor = new RangeListEditor(rangePicker) { Left = fieldX, Top = y, TabIndex = 3 };
        Controls.Add(_inputEditor);
        y += 128;

        AddLabel("Output ranges:", labelX, y + 3);
        _outputEditor = new RangeListEditor(rangePicker) { Left = fieldX, Top = y, TabIndex = 4 };
        Controls.Add(_outputEditor);
        y += 128;

        AddLabel("Keyword args:", labelX, y + 3);
        _kwargsBox = new TextBox
        {
            Left = fieldX,
            Top = y,
            Width = fieldWidth,
            Height = 56,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            AcceptsReturn = true,
            TabIndex = 5,
        };
        Controls.Add(_kwargsBox);
        AddLabel("one name=value per line", fieldX, y + 58, dim: true);
        y += 84;

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
            TabIndex = 6,
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
            TabIndex = 7,
        };
        Controls.Add(cancelButton);

        AcceptButton = saveButton;
        CancelButton = cancelButton;

        if (existing is not null)
        {
            _nameBox.Text = existing.Name;
            _scriptBox.SelectedItem = existing.Script;
            _kwargsBox.Text = KwargsText.Format(existing.Kwargs);
        }
        _inputEditor.LoadFrom(existing?.Input);
        _outputEditor.LoadFrom(existing?.Output);
    }

    private void OnNewScript()
    {
        if (_userScriptsDirectory is null) return;
        try
        {
            var name = TextPromptForm.Prompt(this, "New Script", "Script name:");
            if (string.IsNullOrWhiteSpace(name)) return;
            var fileName = ScriptScaffold.Create(_userScriptsDirectory, name);
            if (!_scriptBox.Items.Contains(fileName)) _scriptBox.Items.Add(fileName);
            _scriptBox.SelectedItem = fileName;
            _errorLabel.Visible = false;
        }
        catch (Exception ex)
        {
            ShowError("Couldn't create the script: " + ex.Message);
        }
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
            _inputEditor.ToBindingText(),
            _outputEditor.ToBindingText(),
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
