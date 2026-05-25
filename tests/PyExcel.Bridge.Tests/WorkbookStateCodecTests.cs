using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class WorkbookStateCodecTests
{
    // -------------------------------------------------------------------------
    // Round-trip
    // -------------------------------------------------------------------------

    [Fact]
    public void Roundtrip_EmptyState_PreservesDefaults()
    {
        var original = WorkbookState.Empty("wb.xlsx");
        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");

        Assert.Equal("wb.xlsx", restored.WorkbookKey);
        Assert.False(restored.Enabled);
        Assert.Null(restored.SelectedScript);
        Assert.Null(restored.PyInput);
        Assert.Null(restored.PyOutput);
        Assert.Null(restored.SelectedActionName);
        Assert.Empty(restored.Actions);
    }

    [Fact]
    public void Roundtrip_FullState_PreservesAllPersistedFields()
    {
        var original = WorkbookState.Empty("wb.xlsx") with
        {
            Enabled = true,
            SelectedScript = "transform.py",
            PyInput = "prices=A1:C10; signals=D1:D10",
            PyOutput = "F1",
            SelectedActionName = "compute",
            Actions = new[]
            {
                new RibbonAction(
                    Name: "compute",
                    Script: "transform.py",
                    Input: "A1:C10",
                    Output: "F1",
                    Kwargs: new Dictionary<string, string> { ["factor"] = "5" }),
            },
        };

        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");

        Assert.True(restored.Enabled);
        Assert.Equal("transform.py", restored.SelectedScript);
        Assert.Equal("prices=A1:C10; signals=D1:D10", restored.PyInput);
        Assert.Equal("F1", restored.PyOutput);
        Assert.Equal("compute", restored.SelectedActionName);
        var a = Assert.Single(restored.Actions);
        Assert.Equal("compute", a.Name);
        Assert.Equal("transform.py", a.Script);
        Assert.Equal("A1:C10", a.Input);
        Assert.Equal("F1", a.Output);
        Assert.NotNull(a.Kwargs);
        Assert.Equal("5", a.Kwargs!["factor"]);
    }

    [Fact]
    public void Roundtrip_ActionWithoutKwargs_KwargsNull()
    {
        var original = WorkbookState.Empty("wb.xlsx") with
        {
            Actions = new[] { new RibbonAction("a", "s.py", "A1", "B1", Kwargs: null) },
        };

        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");

        var a = Assert.Single(restored.Actions);
        Assert.Null(a.Kwargs);
    }

    [Fact]
    public void Roundtrip_MultipleActions_PreservesOrder()
    {
        var original = WorkbookState.Empty("wb.xlsx") with
        {
            Actions = new[]
            {
                new RibbonAction("first", "f.py", "A1", "B1"),
                new RibbonAction("second", "g.py", "A2", "B2"),
                new RibbonAction("third", "h.py", "A3", "B3"),
            },
        };

        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");

        Assert.Equal(3, restored.Actions.Count);
        Assert.Equal(new[] { "first", "second", "third" },
            restored.Actions.Select(x => x.Name).ToArray());
    }

    // -------------------------------------------------------------------------
    // Schema / namespace
    // -------------------------------------------------------------------------

    [Fact]
    public void Serialize_RootCarriesSchemaNamespaceAndVersion()
    {
        var doc = WorkbookStateCodec.Serialize(WorkbookState.Empty("wb.xlsx"));
        Assert.NotNull(doc.Root);
        Assert.Equal(WorkbookStateCodec.XmlNamespace, doc.Root!.Name.NamespaceName);
        Assert.Equal("pyexcel", doc.Root.Name.LocalName);
        Assert.Equal(WorkbookStateCodec.SchemaVersion,
            (string?)doc.Root.Attribute("state-version"));
    }

    [Fact]
    public void Serialize_KwargsEmittedInDeterministicOrder()
    {
        // Different insertion order, same logical kwargs — the on-disk
        // representation must be byte-identical so a no-op save doesn't
        // churn the workbook's binary diff.
        var s1 = WorkbookState.Empty("wb.xlsx") with
        {
            Actions = new[]
            {
                new RibbonAction("a", "s.py", "A1", "B1",
                    new Dictionary<string, string> { ["c"] = "3", ["a"] = "1", ["b"] = "2" }),
            },
        };
        var s2 = WorkbookState.Empty("wb.xlsx") with
        {
            Actions = new[]
            {
                new RibbonAction("a", "s.py", "A1", "B1",
                    new Dictionary<string, string> { ["a"] = "1", ["b"] = "2", ["c"] = "3" }),
            },
        };

        Assert.Equal(
            WorkbookStateCodec.Serialize(s1).ToString(),
            WorkbookStateCodec.Serialize(s2).ToString());
    }

    // -------------------------------------------------------------------------
    // Transient fields are NOT persisted
    // -------------------------------------------------------------------------

    [Fact]
    public void Roundtrip_TransientFields_RestoredToDefaults()
    {
        // CurrentSheet and AvailableScripts are derived from live
        // sources (sheet activate, FileSystemWatcher) and must not be
        // pinned into the workbook — a workbook opened on a fresh
        // machine where the userScripts/ dir is empty should still load
        // cleanly.
        var original = WorkbookState.Empty("wb.xlsx") with
        {
            CurrentSheet = "Sheet1",
            AvailableScripts = new[] { "a.py", "b.py" },
        };

        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");

        Assert.Null(restored.CurrentSheet);
        Assert.Empty(restored.AvailableScripts);
    }

    [Fact]
    public void Roundtrip_WorkbookKey_TakenFromCallerNotDocument()
    {
        // The XML doesn't carry the key — the persister already knows
        // which workbook it's loading from. Different key on
        // Deserialize is intended (e.g. workbook saved as a copy).
        var original = WorkbookState.Empty("wb.xlsx") with { Enabled = true };
        var doc = WorkbookStateCodec.Serialize(original);
        var restored = WorkbookStateCodec.Deserialize(doc, "copy-of-wb.xlsx");

        Assert.Equal("copy-of-wb.xlsx", restored.WorkbookKey);
        Assert.True(restored.Enabled);
    }

    // -------------------------------------------------------------------------
    // Deserialization errors
    // -------------------------------------------------------------------------

    [Fact]
    public void Deserialize_WrongRootElement_Throws()
    {
        var doc = XDocument.Parse(
            $"<other xmlns=\"{WorkbookStateCodec.XmlNamespace}\" state-version=\"1\"/>");
        var ex = Assert.Throws<FormatException>(
            () => WorkbookStateCodec.Deserialize(doc, "wb.xlsx"));
        Assert.Contains("unexpected root", ex.Message);
    }

    [Fact]
    public void Deserialize_WrongNamespace_Throws()
    {
        var doc = XDocument.Parse(
            "<pyexcel xmlns=\"urn:not-us\" state-version=\"1\"><enabled>true</enabled><actions/></pyexcel>");
        Assert.Throws<FormatException>(
            () => WorkbookStateCodec.Deserialize(doc, "wb.xlsx"));
    }

    [Fact]
    public void Deserialize_MissingSchemaVersion_Throws()
    {
        var doc = XDocument.Parse(
            $"<pyexcel xmlns=\"{WorkbookStateCodec.XmlNamespace}\"><enabled>true</enabled><actions/></pyexcel>");
        var ex = Assert.Throws<FormatException>(
            () => WorkbookStateCodec.Deserialize(doc, "wb.xlsx"));
        Assert.Contains("state-version", ex.Message);
    }

    [Fact]
    public void Deserialize_UnsupportedSchemaVersion_Throws()
    {
        var doc = XDocument.Parse(
            $"<pyexcel xmlns=\"{WorkbookStateCodec.XmlNamespace}\" state-version=\"99\"><enabled>true</enabled><actions/></pyexcel>");
        var ex = Assert.Throws<FormatException>(
            () => WorkbookStateCodec.Deserialize(doc, "wb.xlsx"));
        Assert.Contains("99", ex.Message);
    }

    [Fact]
    public void Deserialize_ActionMissingRequiredAttribute_Throws()
    {
        var doc = XDocument.Parse(
            $"<pyexcel xmlns=\"{WorkbookStateCodec.XmlNamespace}\" state-version=\"1\">" +
            "<enabled>true</enabled>" +
            "<actions><action name=\"a\" script=\"s.py\" input=\"A1\"/></actions>" +
            "</pyexcel>");
        var ex = Assert.Throws<FormatException>(
            () => WorkbookStateCodec.Deserialize(doc, "wb.xlsx"));
        Assert.Contains("output", ex.Message);
    }

    [Fact]
    public void Deserialize_AcceptsTrueOneAsEnabled()
    {
        // XmlConvert.ToString emits "true"/"false"; XmlConvert.ToBoolean
        // also accepts "1"/"0" (XML Schema spec). Round-trip via "1"
        // for a hand-edited file should still work.
        var doc = XDocument.Parse(
            $"<pyexcel xmlns=\"{WorkbookStateCodec.XmlNamespace}\" state-version=\"1\">" +
            "<enabled>1</enabled><actions/></pyexcel>");
        var s = WorkbookStateCodec.Deserialize(doc, "wb.xlsx");
        Assert.True(s.Enabled);
    }
}
