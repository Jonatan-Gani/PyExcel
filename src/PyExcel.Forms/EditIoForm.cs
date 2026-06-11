#if NETFRAMEWORK
using System;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.Excel;

namespace PyExcel.Forms;

/// <summary>
/// The Edit-Import / Edit-Export / Edit-Paste dialogs (Phase 8 ports of
/// v1's <c>frmEditImport</c> / <c>frmEditExport</c> / <c>frmEditPaste</c>),
/// unified into one parameterised two-field editor since they differ only
/// in labels and which field browses for a file. Each is a guided way to
/// set the same workbook-state fields the ribbon edit-boxes hold, with
/// inline validation that rejects exactly what the run-time service would.
///
/// <para>Validation runs through the cross-platform
/// <see cref="EditIoValidator"/>, so this class is only the WinForms shell
/// (behind <c>#if NETFRAMEWORK</c>). Shown modally with an owner; OK closes
/// only when the fields validate.</para>
/// </summary>
public sealed class EditIoForm : Form
{
    /// <summary>How a field is edited: a plain range/text box, or a box
    /// with a Browse button opening a file dialog.</summary>
    private enum FieldKind { Range, OpenFile, SaveFile }

    private const string ImportFileFilter =
        "Data files (*.csv;*.tsv;*.txt;*.xlsx;*.xlsm;*.xlsb)|" +
        "*.csv;*.tsv;*.txt;*.xlsx;*.xlsm;*.xlsb|All files (*.*)|*.*";
    private const string ExportFileFilter =
        "CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|All files (*.*)|*.*";

    private readonly TextBox? _inputBox;
    private readonly TextBox _outputBox;
    private readonly Label _errorLabel;
    private readonly Func<string?, string?, EditIoValidationResult> _validate;
    private readonly Func<string?>? _selectionProvider;

    /// <summary>The validated result, valid only after the dialog returns
    /// <see cref="DialogResult.OK"/>.</summary>
    public EditIoValidationResult? Result { get; private set; }

    /// <summary>Show the Edit-Import dialog (source file → target range).</summary>
    public static EditIoValidationResult? PromptImport(
        IWin32Window? owner, string? input, string? output, string? workbookDir,
        Func<string?>? selectionProvider = null)
        => Show(owner, new EditIoForm(
            title: "Edit Import",
            inputLabel: "Source file:", inputKind: FieldKind.OpenFile, input: input,
            outputLabel: "Target range:", outputKind: FieldKind.Range, output: output,
            validate: (i, o) => EditIoValidator.ValidateImport(i, o, workbookDir),
            selectionProvider: selectionProvider));

    /// <summary>Show the Edit-Export dialog (source range → target file).</summary>
    public static EditIoValidationResult? PromptExport(
        IWin32Window? owner, string? input, string? output, string? workbookDir,
        Func<string?>? selectionProvider = null)
        => Show(owner, new EditIoForm(
            title: "Edit Export",
            inputLabel: "Source range:", inputKind: FieldKind.Range, input: input,
            outputLabel: "Target file:", outputKind: FieldKind.SaveFile, output: output,
            validate: (i, o) => EditIoValidator.ValidateExport(i, o, workbookDir),
            selectionProvider: selectionProvider));

    /// <summary>Show the Edit-Paste dialog (target range only).</summary>
    public static EditIoValidationResult? PromptPaste(
        IWin32Window? owner, string? output, Func<string?>? selectionProvider = null)
        => Show(owner, new EditIoForm(
            title: "Edit Paste",
            inputLabel: null, inputKind: FieldKind.Range, input: null,
            outputLabel: "Target range:", outputKind: FieldKind.Range, output: output,
            validate: (i, o) => EditIoValidator.ValidatePaste(o),
            selectionProvider: selectionProvider));

    private static EditIoValidationResult? Show(IWin32Window? owner, EditIoForm form)
    {
        using (form)
        {
            var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
            return result == DialogResult.OK ? form.Result : null;
        }
    }

    private EditIoForm(
        string title,
        string? inputLabel, FieldKind inputKind, string? input,
        string outputLabel, FieldKind outputKind, string? output,
        Func<string?, string?, EditIoValidationResult> validate,
        Func<string?>? selectionProvider)
    {
        _validate = validate;
        _selectionProvider = selectionProvider;

        Text = title;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        Width = 480;

        var y = 14;
        if (inputLabel is not null)
        {
            _inputBox = AddFieldRow(inputLabel, inputKind, input, y, tabStart: 0);
            y += 32;
        }
        _outputBox = AddFieldRow(outputLabel, outputKind, output, y, tabStart: 2);
        y += 38;

        _errorLabel = new Label
        {
            Left = 12,
            Top = y,
            Width = ClientSize.Width - 24,
            Height = 30,
            ForeColor = Color.Firebrick,
            Visible = false,
        };
        Controls.Add(_errorLabel);
        y += 36;

        var okButton = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.None,
            Left = ClientSize.Width - 178,
            Top = y,
            Width = 80,
            TabIndex = 10,
        };
        okButton.Click += OnOkClick;
        Controls.Add(okButton);

        var cancelButton = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = y,
            Width = 80,
            TabIndex = 11,
        };
        Controls.Add(cancelButton);

        AcceptButton = okButton;
        CancelButton = cancelButton;
        ClientSize = new Size(ClientSize.Width, y + 36);
    }

    private TextBox AddFieldRow(string label, FieldKind kind, string? value, int y, int tabStart)
    {
        Controls.Add(new Label
        {
            Text = label,
            Left = 12,
            Top = y + 3,
            AutoSize = true,
        });

        const int fieldX = 100;
        int fieldRight = ClientSize.Width - 12;
        bool isFile = kind != FieldKind.Range;
        bool canPickRange = kind == FieldKind.Range && _selectionProvider is not null;
        bool hasButton = isFile || canPickRange;

        var box = new TextBox
        {
            Left = fieldX,
            Top = y,
            Width = hasButton ? fieldRight - fieldX - 84 : fieldRight - fieldX,
            Text = value ?? string.Empty,
            TabIndex = tabStart,
        };
        Controls.Add(box);

        if (hasButton)
        {
            var button = new Button
            {
                Text = isFile ? "Browse…" : "Pick…",
                Left = fieldRight - 80,
                Top = y - 1,
                Width = 80,
                TabIndex = tabStart + 1,
            };
            if (isFile)
                button.Click += (_, _) => BrowseInto(box, kind);
            else
                button.Click += (_, _) => PickRangeInto(box);
            Controls.Add(button);
        }

        return box;
    }

    private void BrowseInto(TextBox target, FieldKind kind)
    {
        if (kind == FieldKind.OpenFile)
        {
            using var dlg = new OpenFileDialog { Filter = ImportFileFilter, CheckFileExists = true };
            if (!string.IsNullOrWhiteSpace(target.Text)) dlg.FileName = target.Text;
            if (dlg.ShowDialog(this) == DialogResult.OK) target.Text = dlg.FileName;
        }
        else if (kind == FieldKind.SaveFile)
        {
            using var dlg = new SaveFileDialog
            {
                Filter = ExportFileFilter,
                DefaultExt = "csv",
                OverwritePrompt = false, // the export service confirms overwrites itself
            };
            if (!string.IsNullOrWhiteSpace(target.Text)) dlg.FileName = target.Text;
            if (dlg.ShowDialog(this) == DialogResult.OK) target.Text = dlg.FileName;
        }
    }

    private void PickRangeInto(TextBox target)
    {
        var picked = RangePickerForm.Prompt(this, target.Text, _selectionProvider);
        if (picked is not null) target.Text = picked;
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var result = _validate(_inputBox?.Text, _outputBox.Text);
        if (!result.IsValid)
        {
            _errorLabel.Text = result.ErrorMessage;
            _errorLabel.Visible = true;
            return;
        }

        Result = result;
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
