#if NETFRAMEWORK
using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using PyExcel.Excel;

namespace PyExcel.Forms;

/// <summary>
/// The unified Export dialog, run in one of two modes:
/// <list type="bullet">
///   <item><b>Edit</b> (<see cref="PromptDefaults"/>) — configure and save the
///     workbook's <em>default</em> export recipe. Nothing is exported; the
///     primary button is "Save".</item>
///   <item><b>Export</b> (<see cref="PromptExport"/>) — seeded from those saved
///     defaults, lets the user tweak the recipe for this one run, then exports.
///     The primary button is "Export", and a checkbox can promote the tweaked
///     recipe back to the default.</item>
/// </list>
///
/// <para>Both modes edit the same <see cref="ExportSettings"/>: a source range,
/// a destination folder, a base file name, the file type (CSV / TSV), and an
/// optional date/time stamp that keeps every export uniquely named. A live
/// preview shows exactly what the destination file will be called, and (in
/// Export mode) an un-stamped name that already exists prompts before
/// overwriting. The composition rules live in the cross-platform
/// <see cref="ExportSettingsPlanner"/>, so this class is only the WinForms shell
/// (behind <c>#if NETFRAMEWORK</c>).</para>
/// </summary>
public sealed class ExportForm : ScaledForm
{
    private enum Mode { Defaults, Export }

    private const string FolderHint = "(saved next to the workbook)";

    // The timestamp styles offered in the drop-down, parallel to its items.
    private static readonly ExportTimestampStyle[] StampStyles =
    {
        ExportTimestampStyle.DateAndTime,
        ExportTimestampStyle.DateOnly,
        ExportTimestampStyle.Compact,
    };

    private readonly Mode _mode;
    private readonly string? _workbookDir;
    private readonly Func<string?, string?>? _rangePicker;

    private readonly TextBox _sourceBox;
    private readonly TextBox _folderBox;
    private readonly TextBox _nameBox;
    private readonly ComboBox _typeCombo;
    private readonly CheckBox _stampCheck;
    private readonly ComboBox _stampCombo;
    private readonly Label _previewLabel;
    private readonly CheckBox? _saveDefaultCheck;
    private readonly Label _errorLabel;

    /// <summary>The recipe the user built, valid only after the dialog returns
    /// <see cref="DialogResult.OK"/>.</summary>
    public ExportSettings? Result { get; private set; }

    /// <summary>True (Export mode only) when the user asked to also save the
    /// tweaked recipe as the new default.</summary>
    public bool SaveAsDefault => _saveDefaultCheck?.Checked == true;

    /// <summary>Show the <b>Edit</b> dialog — configure and return the recipe to
    /// persist as the workbook's export default, or null on cancel.</summary>
    public static ExportSettings? PromptDefaults(
        IWin32Window? owner, ExportSettings initial, string? workbookDir,
        Func<string?, string?>? rangePicker = null)
    {
        using var form = new ExportForm(Mode.Defaults, initial, workbookDir, rangePicker);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.Result : null;
    }

    /// <summary>Show the <b>Export</b> dialog — seeded from <paramref name="initial"/>
    /// (the saved defaults), return the recipe to export now plus whether to save
    /// it as the new default, or null on cancel.</summary>
    public static ExportPromptResult? PromptExport(
        IWin32Window? owner, ExportSettings initial, string? workbookDir,
        Func<string?, string?>? rangePicker = null)
    {
        using var form = new ExportForm(Mode.Export, initial, workbookDir, rangePicker);
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK && form.Result is not null
            ? new ExportPromptResult(form.Result, form.SaveAsDefault)
            : null;
    }

    private ExportForm(
        Mode mode, ExportSettings initial, string? workbookDir,
        Func<string?, string?>? rangePicker)
    {
        if (initial is null) throw new ArgumentNullException(nameof(initial));
        _mode = mode;
        _workbookDir = workbookDir;
        _rangePicker = rangePicker;

        Text = mode == Mode.Export ? "Export" : "Export Settings";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;

        const int labelX = 14;
        const int fieldX = 118;
        const int formWidth = 488;
        const int fieldRight = formWidth - 14;
        const int buttonWidth = 78;
        ClientSize = new Size(formWidth, 100); // height finalised at the end

        var y = 16;

        AddLabel("Source range:", labelX, y + 3);
        _sourceBox = new TextBox
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX - buttonWidth - 6,
            Text = initial.SourceRange ?? string.Empty, TabIndex = 0,
        };
        Controls.Add(_sourceBox);
        var pick = new Button
        {
            Text = "Pick…", Left = fieldRight - buttonWidth, Top = y - 1,
            Width = buttonWidth, TabIndex = 1, Enabled = _rangePicker is not null,
        };
        pick.Click += (_, _) => PickRangeInto(_sourceBox);
        Controls.Add(pick);
        y += 34;

        AddLabel("Save to folder:", labelX, y + 3);
        _folderBox = new TextBox
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX - buttonWidth - 6,
            Text = initial.Folder ?? string.Empty, TabIndex = 2,
        };
        Controls.Add(_folderBox);
        var browse = new Button
        {
            Text = "Browse…", Left = fieldRight - buttonWidth, Top = y - 1,
            Width = buttonWidth, TabIndex = 3,
        };
        browse.Click += (_, _) => BrowseForFolder();
        Controls.Add(browse);
        y += 34;

        AddLabel("File name:", labelX, y + 3);
        _nameBox = new TextBox
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX,
            Text = initial.BaseName ?? string.Empty, TabIndex = 4,
        };
        Controls.Add(_nameBox);
        y += 34;

        AddLabel("File type:", labelX, y + 3);
        _typeCombo = new ComboBox
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX,
            DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 5,
        };
        _typeCombo.Items.Add(ExportFileType.Csv.Label());
        _typeCombo.Items.Add(ExportFileType.Tsv.Label());
        _typeCombo.SelectedIndex = initial.FileType == ExportFileType.Tsv ? 1 : 0;
        Controls.Add(_typeCombo);
        y += 36;

        _stampCheck = new CheckBox
        {
            Text = "Add a date && time so each export is a new file",
            Left = labelX, Top = y, Width = fieldRight - labelX, Height = 22,
            Checked = initial.Timestamp != ExportTimestampStyle.None, TabIndex = 6,
        };
        Controls.Add(_stampCheck);
        y += 28;

        AddLabel("Stamp format:", labelX, y + 3);
        _stampCombo = new ComboBox
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX,
            DropDownStyle = ComboBoxStyle.DropDownList, TabIndex = 7,
        };
        foreach (var style in StampStyles)
            _stampCombo.Items.Add($"{StyleName(style)}   ({style.Example()})");
        _stampCombo.SelectedIndex = StyleIndex(initial.Timestamp);
        Controls.Add(_stampCombo);
        y += 38;

        AddLabel("Saves as:", labelX, y + 1);
        _previewLabel = new Label
        {
            Left = fieldX, Top = y, Width = fieldRight - fieldX, Height = 34,
            Font = new Font(Font, FontStyle.Bold),
            ForeColor = SystemColors.Highlight,
        };
        Controls.Add(_previewLabel);
        y += 40;

        if (mode == Mode.Export)
        {
            _saveDefaultCheck = new CheckBox
            {
                Text = "Also save these settings as the default",
                Left = labelX, Top = y, Width = fieldRight - labelX, Height = 22,
                Checked = false, TabIndex = 8,
            };
            Controls.Add(_saveDefaultCheck);
            y += 28;
        }

        _errorLabel = new Label
        {
            Left = labelX, Top = y, Width = fieldRight - labelX, Height = 30,
            ForeColor = Color.Firebrick, Visible = false,
        };
        Controls.Add(_errorLabel);
        y += 36;

        var ok = new Button
        {
            Text = mode == Mode.Export ? "Export" : "Save",
            DialogResult = DialogResult.None,
            Left = fieldRight - buttonWidth * 2 - 8, Top = y, Width = buttonWidth, TabIndex = 9,
        };
        ok.Click += OnOkClick;
        Controls.Add(ok);

        var cancel = new Button
        {
            Text = "Cancel", DialogResult = DialogResult.Cancel,
            Left = fieldRight - buttonWidth, Top = y, Width = buttonWidth, TabIndex = 10,
        };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;
        ClientSize = new Size(formWidth, y + 36);

        // Live preview + enable/disable the stamp drop-down with its checkbox.
        _nameBox.TextChanged += (_, _) => UpdatePreview();
        _folderBox.TextChanged += (_, _) => UpdatePreview();
        _typeCombo.SelectedIndexChanged += (_, _) => UpdatePreview();
        _stampCombo.SelectedIndexChanged += (_, _) => UpdatePreview();
        _stampCheck.CheckedChanged += (_, _) => { SyncStampEnabled(); UpdatePreview(); };
        SyncStampEnabled();
        UpdatePreview();
    }

    private void AddLabel(string text, int x, int top)
        => Controls.Add(new Label { Text = text, Left = x, Top = top, AutoSize = true });

    /// <summary>The recipe currently described by the controls.</summary>
    private ExportSettings CurrentSettings()
    {
        var fileType = _typeCombo.SelectedIndex == 1 ? ExportFileType.Tsv : ExportFileType.Csv;
        var timestamp = !_stampCheck.Checked
            ? ExportTimestampStyle.None
            : StampStyles[Math.Max(0, _stampCombo.SelectedIndex)];
        return new ExportSettings(
            _sourceBox.Text, _folderBox.Text, _nameBox.Text, fileType, timestamp);
    }

    /// <summary>Refresh the "Saves as" preview from the live controls, showing the
    /// composed file name (with a real example stamp) and the destination folder.</summary>
    private void UpdatePreview()
    {
        var settings = CurrentSettings();
        var fileName = ExportSettingsPlanner.ComposeFileName(settings, DateTime.Now);
        var folder = string.IsNullOrWhiteSpace(_folderBox.Text) ? FolderHint : _folderBox.Text.Trim();
        _previewLabel.Text = $"{fileName}\n{folder}";
    }

    private void SyncStampEnabled() => _stampCombo.Enabled = _stampCheck.Checked;

    private void PickRangeInto(TextBox target)
    {
        if (_rangePicker is null) return;
        var picked = RangePick.OnSheet(_rangePicker, target.Text, this);
        if (picked is not null) target.Text = picked;
    }

    private void BrowseForFolder()
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Choose the folder to export into.",
            ShowNewFolderButton = true,
        };
        var start = _folderBox.Text.Trim();
        if (start.Length == 0) start = _workbookDir ?? string.Empty;
        if (start.Length > 0 && Directory.Exists(start)) dlg.SelectedPath = start;
        if (dlg.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
            _folderBox.Text = dlg.SelectedPath;
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var settings = CurrentSettings();

        if (_mode == Mode.Export)
        {
            // Resolve validates the source range and gives us the concrete path.
            ExportPlan plan;
            try
            {
                plan = ExportSettingsPlanner.Resolve(settings, DateTime.Now, _workbookDir);
            }
            catch (FormatException ex)
            {
                ShowError(ex.Message);
                return;
            }

            // An un-stamped name is deterministic, so it can clobber an existing
            // file — confirm first. A stamped name is unique by construction.
            if (settings.Timestamp == ExportTimestampStyle.None && File.Exists(plan.AbsoluteTargetPath))
            {
                var answer = MessageBox.Show(
                    this,
                    $"'{Path.GetFileName(plan.AbsoluteTargetPath)}' already exists in that folder.\n\n" +
                    "Overwrite it?",
                    "Export",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (answer != DialogResult.Yes) return;
            }
        }

        Result = settings;
        DialogResult = DialogResult.OK;
        Close();
    }

    private void ShowError(string message)
    {
        _errorLabel.Text = message;
        _errorLabel.Visible = true;
    }

    // Drop-down item text. ComboBox items aren't mnemonic-processed, so a literal
    // ampersand would render doubled — spell out "and" instead.
    private static string StyleName(ExportTimestampStyle style) => style switch
    {
        ExportTimestampStyle.DateAndTime => "Date and time",
        ExportTimestampStyle.DateOnly => "Date only",
        ExportTimestampStyle.Compact => "Compact",
        _ => "None",
    };

    /// <summary>The drop-down index for a style, defaulting to the first
    /// (date &amp; time) for <see cref="ExportTimestampStyle.None"/> so toggling the
    /// checkbox on lands on a sensible choice.</summary>
    private static int StyleIndex(ExportTimestampStyle style)
    {
        var idx = Array.IndexOf(StampStyles, style);
        return idx < 0 ? 0 : idx;
    }
}
#endif
