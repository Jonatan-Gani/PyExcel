#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PyExcel.Excel;

namespace PyExcel.Forms;

/// <summary>
/// The Export Wizard (Phase 8 port of v1's <c>frmExportWizard</c>, reshaped
/// to v2's CSV/TSV export): a multi-row editor mapping source ranges to
/// target files, run in one batch. v1's dead Add/Remove row buttons are
/// fixed by construction — rows live in a <see cref="FlowLayoutPanel"/> that
/// reflows automatically on add/remove, so there's no manual repositioning
/// to get wrong.
///
/// <para>Validation runs through the cross-platform
/// <see cref="ExportBatchValidator"/> (which reuses <see cref="ExportPlanner"/>
/// per row); OK closes only when every row is valid.</para>
/// </summary>
public sealed class ExportWizardForm : ScaledForm
{
    private const string ExportFileFilter =
        "CSV (*.csv)|*.csv|TSV (*.tsv)|*.tsv|All files (*.*)|*.*";
    private const int RowWidth = 540;

    private readonly FlowLayoutPanel _rows;
    private readonly Label _errorLabel;
    private readonly Func<IReadOnlyList<ExportJob>, ExportBatchValidationResult> _validate;

    public IReadOnlyList<ExportJob>? Result { get; private set; }

    /// <summary>Show the wizard and return the validated jobs, or null on
    /// cancel.</summary>
    public static IReadOnlyList<ExportJob>? Prompt(
        IWin32Window? owner, IReadOnlyList<ExportJob>? initial, string? workbookDir)
    {
        using var form = new ExportWizardForm(
            initial, jobs => ExportBatchValidator.Validate(jobs, workbookDir));
        var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
        return result == DialogResult.OK ? form.Result : null;
    }

    private ExportWizardForm(
        IReadOnlyList<ExportJob>? initial,
        Func<IReadOnlyList<ExportJob>, ExportBatchValidationResult> validate)
    {
        _validate = validate;

        Text = "Export Wizard";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        Font = SystemFonts.MessageBoxFont;
        ClientSize = new Size(580, 340);

        Controls.Add(new Label { Text = "Source range", Left = 16, Top = 10, AutoSize = true });
        Controls.Add(new Label { Text = "Target file", Left = 160, Top = 10, AutoSize = true });

        _rows = new FlowLayoutPanel
        {
            Left = 12,
            Top = 30,
            Width = RowWidth + 24,
            Height = 220,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            BorderStyle = BorderStyle.FixedSingle,
        };
        Controls.Add(_rows);

        var addRow = new Button
        {
            Text = "Add row",
            Left = 12,
            Top = 258,
            Width = 90,
            TabIndex = 100,
        };
        addRow.Click += (_, _) => AddRow(string.Empty, string.Empty);
        Controls.Add(addRow);

        _errorLabel = new Label
        {
            Left = 12,
            Top = 290,
            Width = ClientSize.Width - 24,
            Height = 16,
            ForeColor = Color.Firebrick,
            Visible = false,
        };
        Controls.Add(_errorLabel);

        var ok = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.None,
            Left = ClientSize.Width - 178,
            Top = ClientSize.Height - 32,
            Width = 80,
            TabIndex = 101,
        };
        ok.Click += OnOkClick;
        Controls.Add(ok);

        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Left = ClientSize.Width - 92,
            Top = ClientSize.Height - 32,
            Width = 80,
            TabIndex = 102,
        };
        Controls.Add(cancel);

        AcceptButton = ok;
        CancelButton = cancel;

        // Seed rows (at least one so the dialog is never empty).
        if (initial is { Count: > 0 })
            foreach (var job in initial) AddRow(job.SourceRange, job.TargetPath);
        else
            AddRow(string.Empty, string.Empty);
    }

    private void AddRow(string source, string target)
    {
        var panel = new Panel { Width = RowWidth, Height = 30, Margin = new Padding(0, 0, 0, 2) };

        var sourceBox = new TextBox { Left = 4, Top = 4, Width = 140, Text = source };
        var targetBox = new TextBox { Left = 150, Top = 4, Width = 236, Text = target };

        var browse = new Button { Text = "Browse…", Left = 392, Top = 3, Width = 70 };
        browse.Click += (_, _) =>
        {
            using var dlg = new SaveFileDialog
            {
                Filter = ExportFileFilter,
                DefaultExt = "csv",
                OverwritePrompt = false,
            };
            if (!string.IsNullOrWhiteSpace(targetBox.Text)) dlg.FileName = targetBox.Text;
            if (dlg.ShowDialog(this) == DialogResult.OK) targetBox.Text = dlg.FileName;
        };

        var remove = new Button { Text = "✕", Left = 468, Top = 3, Width = 28 };
        remove.Click += (_, _) =>
        {
            _rows.Controls.Remove(panel);
            panel.Dispose();
            // Never leave the wizard with zero rows.
            if (_rows.Controls.Count == 0) AddRow(string.Empty, string.Empty);
        };

        panel.Controls.Add(sourceBox);
        panel.Controls.Add(targetBox);
        panel.Controls.Add(browse);
        panel.Controls.Add(remove);
        _rows.Controls.Add(panel);
        // Rows added after the form has loaded must be brought up to the same
        // DPI scale as the rest of the layout (no-op during construction — the
        // initial rows are scaled by the base form's load-time pass).
        ScaleNewControl(panel);
    }

    private void OnOkClick(object? sender, EventArgs e)
    {
        var jobs = new List<ExportJob>();
        foreach (Control panel in _rows.Controls)
        {
            string src = string.Empty, tgt = string.Empty;
            foreach (Control c in panel.Controls)
            {
                if (c is TextBox tb)
                {
                    if (tb.Left < 150) src = tb.Text;
                    else tgt = tb.Text;
                }
            }
            jobs.Add(new ExportJob(src, tgt));
        }

        var result = _validate(jobs);
        if (!result.IsValid)
        {
            _errorLabel.Text = result.ErrorMessage;
            _errorLabel.Visible = true;
            return;
        }

        Result = result.Jobs;
        DialogResult = DialogResult.OK;
        Close();
    }
}
#endif
