#if NETFRAMEWORK
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// A reusable list editor for an ordered set of range bindings (Note 4): a
/// ListBox plus Add / Edit / Remove / Up / Down, where Add and Edit open
/// <see cref="RangeNameForm"/> (a native range pick + optional name). The
/// EditAction dialog hosts two of these — one for inputs, one for outputs —
/// in place of the old single text field. Loads from and serialises back to
/// the ribbon's <c>{name}=range; …</c> syntax via <see cref="RibbonRangeParser"/>.
///
/// <para>Fixed 316×120 layout so the host only sets its location; the
/// cross-platform parse/format it delegates to is unit-tested.</para>
/// </summary>
internal sealed class RangeListEditor : UserControl
{
    private readonly ListBox _list;
    private readonly List<RangeBinding> _items = new();
    private readonly Func<string?, string?>? _rangePicker;
    private readonly Button _editButton;
    private readonly Button _removeButton;
    private readonly Button _upButton;
    private readonly Button _downButton;

    public RangeListEditor(Func<string?, string?>? rangePicker)
    {
        _rangePicker = rangePicker;

        Width = 316;
        Height = 120;
        const int btnW = 72;
        const int btnH = 22;
        const int gap = 6;

        _list = new ListBox
        {
            Left = 0,
            Top = 0,
            Width = Width - btnW - gap,
            Height = Height,
            IntegralHeight = false,
        };
        _list.SelectedIndexChanged += (_, _) => UpdateButtons();
        _list.DoubleClick += (_, _) => EditSelected();
        Controls.Add(_list);

        int bx = Width - btnW;
        AddButton("Add", bx, 0, btnW, btnH, (_, _) => AddNew());
        _editButton = AddButton("Edit", bx, 24, btnW, btnH, (_, _) => EditSelected());
        _removeButton = AddButton("Remove", bx, 48, btnW, btnH, (_, _) => RemoveSelected());
        _upButton = AddButton("Up", bx, 72, btnW, btnH, (_, _) => MoveItem(-1));
        _downButton = AddButton("Down", bx, 96, btnW, btnH, (_, _) => MoveItem(1));

        UpdateButtons();
    }

    private Button AddButton(string text, int x, int y, int w, int h, EventHandler onClick)
    {
        var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h };
        b.Click += onClick;
        Controls.Add(b);
        return b;
    }

    /// <summary>Parse <paramref name="bindingText"/> (the ribbon syntax) into
    /// the list. Malformed text leaves the list empty — the field was free
    /// text before, so this never throws.</summary>
    public void LoadFrom(string? bindingText)
    {
        _items.Clear();
        try
        {
            foreach (var b in RibbonRangeParser.Parse(bindingText)) _items.Add(b);
        }
        catch (FormatException)
        {
            _items.Clear();
        }
        Rebuild();
    }

    /// <summary>Serialise the list back to the ribbon's {name}=range syntax.</summary>
    public string ToBindingText() => RibbonRangeParser.Format(_items);

    private void AddNew()
    {
        var b = RangeNameForm.Prompt(FindForm(), null, _rangePicker);
        if (b is null) return;
        _items.Add(b);
        Rebuild();
        _list.SelectedIndex = _items.Count - 1;
    }

    private void EditSelected()
    {
        var i = _list.SelectedIndex;
        if (i < 0) return;
        var b = RangeNameForm.Prompt(FindForm(), _items[i], _rangePicker);
        if (b is null) return;
        _items[i] = b;
        Rebuild();
        _list.SelectedIndex = i;
    }

    private void RemoveSelected()
    {
        var i = _list.SelectedIndex;
        if (i < 0) return;
        _items.RemoveAt(i);
        Rebuild();
        if (_items.Count > 0) _list.SelectedIndex = Math.Min(i, _items.Count - 1);
    }

    private void MoveItem(int delta)
    {
        var i = _list.SelectedIndex;
        var j = i + delta;
        if (i < 0 || j < 0 || j >= _items.Count) return;
        (_items[j], _items[i]) = (_items[i], _items[j]);
        Rebuild();
        _list.SelectedIndex = j;
    }

    private void Rebuild()
    {
        _list.BeginUpdate();
        _list.Items.Clear();
        foreach (var b in _items) _list.Items.Add(Describe(b));
        _list.EndUpdate();
        UpdateButtons();
    }

    private static string Describe(RangeBinding b)
        => string.IsNullOrEmpty(b.Name) ? b.RangeText : b.Name + " = " + b.RangeText;

    private void UpdateButtons()
    {
        var i = _list.SelectedIndex;
        var has = i >= 0;
        _editButton.Enabled = has;
        _removeButton.Enabled = has;
        _upButton.Enabled = has && i > 0;
        _downButton.Enabled = has && i < _items.Count - 1;
    }
}
#endif
