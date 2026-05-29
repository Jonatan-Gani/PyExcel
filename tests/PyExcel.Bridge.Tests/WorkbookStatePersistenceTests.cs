using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Tests for the string-level persistence helpers the Windows-only
/// <c>WorkbookStatePersister</c> delegates to:
/// <see cref="WorkbookStateCodec.SerializeToString"/> and
/// <see cref="WorkbookStateCodec.TryDeserialize"/>. The COM read/write of
/// the <c>CustomXMLPart</c> itself can't run on Linux, but everything that
/// can go wrong in the XML layer is covered here.
/// </summary>
public class WorkbookStatePersistenceTests
{
    private static WorkbookState Sample(string key) =>
        WorkbookState.Empty(key) with
        {
            Enabled = true,
            SelectedScript = "transform.py",
            PyInput = "A1:B2",
            PyOutput = "D1",
            Actions = new[] { new RibbonAction("act", "a.py", "A1", "B1") },
            SelectedActionName = "act",
        };

    [Fact]
    public void SerializeToString_ThenTryDeserialize_RoundTrips()
    {
        var original = Sample(@"C:\wb.xlsx");
        string xml = WorkbookStateCodec.SerializeToString(original);

        Assert.True(WorkbookStateCodec.TryDeserialize(xml, @"C:\wb.xlsx", out var restored));
        Assert.NotNull(restored);
        Assert.True(restored!.Enabled);
        Assert.Equal("transform.py", restored.SelectedScript);
        Assert.Equal("A1:B2", restored.PyInput);
        Assert.Equal("D1", restored.PyOutput);
        Assert.Equal("act", restored.SelectedActionName);
        Assert.Single(restored.Actions);
        Assert.Equal("a.py", restored.Actions[0].Script);
    }

    [Fact]
    public void SerializeToString_CarriesTheLocatorNamespace()
    {
        // The COM persister finds PyExcel's part via SelectByNamespace; if
        // the serialised string ever dropped the namespace, the part would
        // be unfindable on reopen.
        string xml = WorkbookStateCodec.SerializeToString(WorkbookState.Empty("k"));
        Assert.Contains(WorkbookStateCodec.XmlNamespace, xml);
    }

    [Fact]
    public void TryDeserialize_KeysByTheCallerSuppliedKey()
    {
        // A workbook saved-as-a-copy gets a new key; the persisted XML never
        // carried the key, so the caller's key must win.
        string xml = WorkbookStateCodec.SerializeToString(Sample("old-key"));
        Assert.True(WorkbookStateCodec.TryDeserialize(xml, "new-key", out var restored));
        Assert.Equal("new-key", restored!.WorkbookKey);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void TryDeserialize_NullOrBlank_ReturnsFalse(string? xml)
    {
        Assert.False(WorkbookStateCodec.TryDeserialize(xml, "k", out var state));
        Assert.Null(state);
    }

    [Fact]
    public void TryDeserialize_NotXml_ReturnsFalse()
    {
        Assert.False(WorkbookStateCodec.TryDeserialize("this is not xml <<<", "k", out var state));
        Assert.Null(state);
    }

    [Fact]
    public void TryDeserialize_ForeignNamespace_ReturnsFalse()
    {
        // A non-PyExcel CustomXMLPart that happens to share the collection.
        const string foreign = "<root xmlns=\"urn:someone:else\"><x/></root>";
        Assert.False(WorkbookStateCodec.TryDeserialize(foreign, "k", out var state));
        Assert.Null(state);
    }

    [Fact]
    public void TryDeserialize_WrongSchemaVersion_ReturnsFalse()
    {
        const string wrong =
            "<pyexcel state-version=\"99\" xmlns=\"urn:pyexcel:state:1\">" +
            "<enabled>true</enabled></pyexcel>";
        Assert.False(WorkbookStateCodec.TryDeserialize(wrong, "k", out var state));
        Assert.Null(state);
    }
}
