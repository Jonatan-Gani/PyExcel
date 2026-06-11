using System;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class LegacyStateConverterTests
{
    private const string GS = "\u001D"; // v1 action separator (Chr(29))

    // -------------------------------------------------------------------------
    // ParseActions — the serialized v1 "Actions" Name value
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ParseActions_NullOrEmpty_ReturnsEmpty(string? raw)
    {
        Assert.Empty(LegacyStateConverter.ParseActions(raw));
    }

    [Fact]
    public void ParseActions_NamedFormat_SingleAction()
    {
        var raw = "compute|script=transform.py|input=A1:C10|output=F1|entireRow=False|entreToEnd=False" + GS;

        var actions = LegacyStateConverter.ParseActions(raw);

        var a = Assert.Single(actions);
        Assert.Equal("compute", a.Name);
        Assert.Equal("transform.py", a.Script);
        Assert.Equal("A1:C10", a.Input);
        Assert.Equal("F1", a.Output);
        Assert.Null(a.Kwargs);
    }

    [Fact]
    public void ParseActions_NamedFormat_MultipleActions()
    {
        var raw =
            "first|script=a.py|input=A1|output=B1" + GS +
            "second|script=b.py|input=C1|output=D1" + GS;

        var actions = LegacyStateConverter.ParseActions(raw);

        Assert.Equal(2, actions.Count);
        Assert.Equal("first", actions[0].Name);
        Assert.Equal("second", actions[1].Name);
        Assert.Equal("b.py", actions[1].Script);
    }

    [Fact]
    public void ParseActions_RepeatedInputOutput_AccumulatesWithSemicolonSpace()
    {
        var raw = "a|script=s.py|input=A1:A5|input=C1:C5|output=E1|output=F1" + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));

        Assert.Equal("A1:A5; C1:C5", a.Input);
        Assert.Equal("E1; F1", a.Output);
    }

    [Fact]
    public void ParseActions_LegacyPositionalFormat()
    {
        // No '=' in the second token → positional: name|script|input|output|er|ete
        var raw = "legacy|old.py|A1:B2|C1|True|False" + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));

        Assert.Equal("legacy", a.Name);
        Assert.Equal("old.py", a.Script);
        Assert.Equal("A1:B2", a.Input);
        Assert.Equal("C1", a.Output);
    }

    [Fact]
    public void ParseActions_LineFeedSeparatorFallback()
    {
        var raw = "a|script=x.py|input=A1|output=B1\nb|script=y.py|input=C1|output=D1";

        var actions = LegacyStateConverter.ParseActions(raw);

        Assert.Equal(2, actions.Count);
        Assert.Equal("a", actions[0].Name);
        Assert.Equal("b", actions[1].Name);
    }

    [Fact]
    public void ParseActions_SemicolonSeparatorFallback()
    {
        // Oldest format: ';' between positional actions, no '=' and no Chr(29)/Chr(10).
        var raw = "a|x.py|A1|B1;b|y.py|C1|D1";

        var actions = LegacyStateConverter.ParseActions(raw);

        Assert.Equal(2, actions.Count);
        Assert.Equal("x.py", actions[0].Script);
        Assert.Equal("y.py", actions[1].Script);
    }

    [Fact]
    public void ParseActions_DropsEntireRowAndEntreToEndFields()
    {
        var raw = "a|script=s.py|input=A1|output=B1|entireRow=True|entreToEnd=True" + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));

        // No v2 home for these — only the four real fields survive.
        Assert.Equal("a", a.Name);
        Assert.Equal("s.py", a.Script);
        Assert.Equal("A1", a.Input);
        Assert.Equal("B1", a.Output);
    }

    [Fact]
    public void ParseActions_DuplicateNames_KeepsFirst()
    {
        var raw =
            "dup|script=first.py|input=A1|output=B1" + GS +
            "dup|script=second.py|input=C1|output=D1" + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));
        Assert.Equal("first.py", a.Script);
    }

    [Fact]
    public void ParseActions_SkipsBlankRows()
    {
        var raw = GS + "a|script=s.py|input=A1|output=B1" + GS + "   " + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));
        Assert.Equal("a", a.Name);
    }

    [Fact]
    public void ParseActions_ValueContainingEquals_SplitsOnFirstOnly()
    {
        var raw = "a|script=s.py|input=A1|output=B1|note=k=v=w" + GS;

        // 'note' is an unknown key → dropped, but the split-on-first behaviour
        // must not throw or mis-parse the real fields around it.
        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));
        Assert.Equal("s.py", a.Script);
        Assert.Equal("A1", a.Input);
    }

    [Fact]
    public void ParseActions_MalformedRowWithNoFields_Skipped()
    {
        // A bare name with no pipe/field (UBound(cols) < 1 in v1) is ignored.
        var raw = "justaname" + GS + "real|script=s.py|input=A1|output=B1" + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));
        Assert.Equal("real", a.Name);
    }

    [Fact]
    public void ParseActions_TrimsNameAndFields()
    {
        var raw = "  spaced  |script=  s.py  |input=  A1  |output=  B1  " + GS;

        var a = Assert.Single(LegacyStateConverter.ParseActions(raw));
        Assert.Equal("spaced", a.Name);
        Assert.Equal("s.py", a.Script);
        Assert.Equal("A1", a.Input);
        Assert.Equal("B1", a.Output);
    }

    // -------------------------------------------------------------------------
    // Convert — the full per-sheet v1 record → WorkbookState
    // -------------------------------------------------------------------------

    [Fact]
    public void Convert_FullState_MapsEveryField()
    {
        var legacy = new LegacyWorkbookState
        {
            Enabled = "1",
            SelectedAction = "compute",
            Actions = "compute|script=transform.py|input=A1:C10|output=F1" + GS,
            SelectedScript = "transform.py",
            PyInput = "A1:C10; D1:D10",
            PyOutput = "F1",
            ImportInput = @"C:\data\in.csv",
            ImportOutput = "A1",
            ExportInput = "A1:C10",
            ExportOutput = @"C:\data\out.csv",
            PasteOutput = "B2",
        };

        var state = LegacyStateConverter.Convert(legacy, "wb.xlsx");

        Assert.Equal("wb.xlsx", state.WorkbookKey);
        Assert.True(state.Enabled);
        Assert.Equal("compute", state.SelectedActionName);
        Assert.Equal("transform.py", state.SelectedScript);
        Assert.Equal("A1:C10; D1:D10", state.PyInput);
        Assert.Equal("F1", state.PyOutput);
        Assert.Equal(@"C:\data\in.csv", state.ImportInput);
        Assert.Equal("A1", state.ImportOutput);
        Assert.Equal("A1:C10", state.ExportInput);
        Assert.Equal(@"C:\data\out.csv", state.ExportOutput);
        Assert.Equal("B2", state.PasteOutput);

        var a = Assert.Single(state.Actions);
        Assert.Equal("compute", a.Name);
        Assert.NotNull(state.SelectedAction);
        Assert.Equal("transform.py", state.SelectedAction!.Script);
    }

    [Fact]
    public void Convert_BlankTextFields_BecomeNull()
    {
        var legacy = new LegacyWorkbookState
        {
            Enabled = "0",
            SelectedScript = "",
            PyInput = "   ",
            PyOutput = null,
            ImportInput = "",
        };

        var state = LegacyStateConverter.Convert(legacy, "wb.xlsx");

        Assert.False(state.Enabled);
        Assert.Null(state.SelectedScript);
        Assert.Null(state.PyInput);
        Assert.Null(state.PyOutput);
        Assert.Null(state.ImportInput);
        Assert.Empty(state.Actions);
        Assert.Null(state.SelectedActionName);
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData("0", false)]
    [InlineData("true", true)]
    [InlineData("True", true)]
    [InlineData("false", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    [InlineData("  1  ", true)]
    public void Convert_EnabledFlag_ParsesV1Encodings(string? raw, bool expected)
    {
        var state = LegacyStateConverter.Convert(new LegacyWorkbookState { Enabled = raw }, "wb.xlsx");
        Assert.Equal(expected, state.Enabled);
    }

    [Fact]
    public void Convert_TrimsScalarFields()
    {
        var legacy = new LegacyWorkbookState
        {
            SelectedScript = "  s.py  ",
            PyOutput = "  F1  ",
            SelectedAction = "  act  ",
        };

        var state = LegacyStateConverter.Convert(legacy, "wb.xlsx");

        Assert.Equal("s.py", state.SelectedScript);
        Assert.Equal("F1", state.PyOutput);
        Assert.Equal("act", state.SelectedActionName);
    }

    [Fact]
    public void Convert_EmptyLegacy_EqualsEmptyStateWithKey()
    {
        var state = LegacyStateConverter.Convert(new LegacyWorkbookState(), "wb.xlsx");

        Assert.Equal(WorkbookState.Empty("wb.xlsx"), state);
    }

    [Fact]
    public void Convert_NullLegacy_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => LegacyStateConverter.Convert(null!, "wb.xlsx"));
    }

    [Fact]
    public void Convert_NullWorkbookKey_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => LegacyStateConverter.Convert(new LegacyWorkbookState(), null!));
    }

    // -------------------------------------------------------------------------
    // The migration end-to-end: v1 record → WorkbookState → codec XML → state
    // -------------------------------------------------------------------------

    [Fact]
    public void Convert_ThenRoundTripsThroughCodec()
    {
        var legacy = new LegacyWorkbookState
        {
            Enabled = "1",
            SelectedAction = "compute",
            Actions =
                "compute|script=transform.py|input=A1:C10|input=D1:D10|output=F1" + GS +
                "other|script=p.py|input=G1|output=H1" + GS,
            SelectedScript = "transform.py",
            PyInput = "A1:C10",
            PyOutput = "F1",
            ExportOutput = @"C:\out.csv",
        };

        var migrated = LegacyStateConverter.Convert(legacy, "wb.xlsx");

        // The whole point: what the converter produces must survive the exact
        // serializer the COM persister will use to write the CustomXMLPart.
        var xml = WorkbookStateCodec.SerializeToString(migrated);
        Assert.True(WorkbookStateCodec.TryDeserialize(xml, "wb.xlsx", out var restored));

        // Field-by-field: WorkbookState's record equality falls back to
        // reference equality on the Actions list, which a codec round-trip
        // necessarily rebuilds — so compare the persisted slice directly.
        Assert.True(restored!.Enabled);
        Assert.Equal("compute", restored.SelectedActionName);
        Assert.Equal("transform.py", restored.SelectedScript);
        Assert.Equal("A1:C10", restored.PyInput);
        Assert.Equal("F1", restored.PyOutput);
        Assert.Equal(@"C:\out.csv", restored.ExportOutput);
        Assert.Equal(2, restored.Actions.Count);
        Assert.Equal("A1:C10; D1:D10", restored.Actions[0].Input);
        // RibbonAction records (Kwargs null) compare by value here.
        Assert.Equal(migrated.Actions[0], restored.Actions[0]);
        Assert.Equal(migrated.Actions[1], restored.Actions[1]);
    }
}
