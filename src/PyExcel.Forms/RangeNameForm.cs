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
internal sealed class RangeNameForm : ScaledForm
{
    private readonly TextBox _rangeBox;
    private readonly TextBox _nameBox;
    private readonly ComboBox _typeBox;
    private readonly Label _typeHint;
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
        ClientSize = new Size(380, 212);

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
            Text = "optional — leave blank to be auto-named by type",
            Left = 80,
            Top = 70,
            AutoSize = true,
            ForeColor = SystemColors.GrayText,
        });

        Controls.Add(new Label { Text = "Type:", Left = 12, Top = 97, AutoSize = true });
        _typeBox = new ComboBox
        {
            Left = 80,
            Top = 94,
            Width = 288,
            DropDownStyle = ComboBoxStyle.DropDownList,
            TabIndex = 3,
        };
        foreach (var t in PyExcelTypes.All)
            _typeBox.Items.Add(PyExcelTypes.DisplayName(t));
        _typeBox.SelectedIndex = IndexOf(initial?.DeclaredType ?? PyExcelType.Auto);
        Controls.Add(_typeBox);

        _typeHint = new Label
        {
            Left = 80,
            Top = 118,
            Width = 288,
            AutoSize = false,
            Height = 16,
            ForeColor = SystemColors.GrayText,
        };
        Controls.Add(_typeHint);
        _typeBox.SelectedIndexChanged += (_, _) => UpdateTypeHint();
        UpdateTypeHint();

        _errorLabel = new Label
        {
            Left = 12,
            Top = 140,
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
            TabIndex = 4,
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
            TabIndex = 5,
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
        Result = new RangeBinding(
            name.Length == 0 ? null : name, result.Address!, SelectedType());
        DialogResult = DialogResult.OK;
        Close();
    }

    /// <summary>The type currently chosen in the box.</summary>
    private PyExcelType SelectedType()
    {
        var index = _typeBox.SelectedIndex;
        return index >= 0 && index < PyExcelTypes.All.Count
            ? PyExcelTypes.All[index]
            : PyExcelType.Auto;
    }

    /// <summary>Position of a type in the dropdown, defaulting to Auto.</summary>
    private static int IndexOf(PyExcelType type)
    {
        for (var i = 0; i < PyExcelTypes.All.Count; i++)
            if (PyExcelTypes.All[i] == type) return i;
        return 0;
    }

    /// <summary>
    /// One line under the box saying what the chosen type will actually do
    /// with the selected cells. The dialog only ever holds an address
    /// string — it has not measured the range — so this describes the rule
    /// rather than predicting the result for this particular selection.
    /// </summary>
    private void UpdateTypeHint() => _typeHint.Text = SelectedType() switch
    {
        PyExcelType.Auto => "by size: a block → DataFrame, one row/column → List, one cell → Scalar",
        PyExcelType.DataFrame => "first row becomes the column headers",
        PyExcelType.Series => "one row or column only; the first cell names it",
        PyExcelType.List => "every cell; a block becomes a list of rows",
        PyExcelType.Tuple => "every cell; a block becomes a tuple of rows",
        PyExcelType.Set => "the distinct cell values",
        PyExcelType.Dict => "2 columns → key/value; 3+ → lists keyed by the header row",
        PyExcelType.NDArray => "a numpy array shaped like the range",
        PyExcelType.Scalar => "a single cell only",
        _ => string.Empty,
    };
}
#endif
