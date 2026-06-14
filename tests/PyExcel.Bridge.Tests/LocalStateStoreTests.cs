using System;
using System.IO;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers <see cref="LocalStateStore"/> — the reliable per-user persistence
/// that makes an enabled workbook come back enabled on reopen without
/// depending on Excel's save / CustomXMLPart round-trip.
/// </summary>
public class LocalStateStoreTests
{
    [Fact]
    public void Save_Then_Load_RoundTrips_Enabled_ProjectDir_And_Actions()
    {
        var key = SavedKey();
        try
        {
            var state = WorkbookState.Empty(key) with
            {
                Enabled = true,
                ProjectDir = "/some/project",
                Actions = new[] { new RibbonAction("a", "s.py", "in", "out") },
            };

            LocalStateStore.Save(key, state);
            var loaded = LocalStateStore.TryLoad(key);

            Assert.NotNull(loaded);
            Assert.True(loaded!.Enabled);
            Assert.Equal("/some/project", loaded.ProjectDir);
            Assert.Single(loaded.Actions);
            Assert.Equal("a", loaded.Actions[0].Name);
        }
        finally { LocalStateStore.Remove(key); }
    }

    [Fact]
    public void Unsaved_Workbook_Keys_Are_Not_Persisted()
    {
        var key = WorkbookKeys.UnsavedKey("Book1");
        LocalStateStore.Save(key, WorkbookState.Empty(key) with { Enabled = true });
        Assert.Null(LocalStateStore.TryLoad(key));
    }

    [Fact]
    public void TryLoad_Unknown_Key_Returns_Null()
        => Assert.Null(LocalStateStore.TryLoad(SavedKey()));

    [Fact]
    public void Remove_Deletes_The_Stored_State()
    {
        var key = SavedKey();
        LocalStateStore.Save(key, WorkbookState.Empty(key) with { Enabled = true });
        Assert.NotNull(LocalStateStore.TryLoad(key));

        LocalStateStore.Remove(key);
        Assert.Null(LocalStateStore.TryLoad(key));
    }

    // A realistic "saved workbook" key is its full path; use a unique temp path
    // so tests don't collide and never touch a real workbook's stored state.
    private static string SavedKey()
        => Path.Combine(Path.GetTempPath(), "pyexcel-lss-" + Guid.NewGuid().ToString("N") + ".xlsx");
}
