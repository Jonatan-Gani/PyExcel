using System.Collections.Generic;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// The per-sheet dimension of <see cref="StateService"/>: sheet-scoped fields
/// (script, Run/Import/Export/Paste bindings, actions) switch with the current
/// sheet, while workbook-scoped facts (enabled, project dir) are shared; plus the
/// persistence bridge (<see cref="StateService.GetProfile"/> /
/// <see cref="StateService.LoadProfile"/>) and default-bucket inheritance.
/// </summary>
public class PerSheetStateTests
{
    [Fact]
    public void SheetScopedFields_AreIsolatedPerSheet()
    {
        var svc = new StateService();
        svc.SetCurrentSheet("wb", "Sheet1");
        svc.SetPyInput("wb", "A1");

        svc.SetCurrentSheet("wb", "Sheet2");
        Assert.Null(svc.Get("wb").PyInput);          // a fresh sheet starts blank
        svc.SetPyInput("wb", "Z9");

        svc.SetCurrentSheet("wb", "Sheet1");
        Assert.Equal("A1", svc.Get("wb").PyInput);   // Sheet1 kept its own value

        svc.SetCurrentSheet("wb", "Sheet2");
        Assert.Equal("Z9", svc.Get("wb").PyInput);
    }

    [Fact]
    public void Actions_AreIsolatedPerSheet()
    {
        var svc = new StateService();
        svc.SetCurrentSheet("wb", "Sheet1");
        svc.AddAction("wb", new RibbonAction("one", "a.py", "A1", "B1"));

        svc.SetCurrentSheet("wb", "Sheet2");
        Assert.Empty(svc.Get("wb").Actions);
        svc.AddAction("wb", new RibbonAction("two", "b.py", "C1", "D1"));

        svc.SetCurrentSheet("wb", "Sheet1");
        var s1 = svc.Get("wb");
        Assert.Single(s1.Actions);
        Assert.Equal("one", s1.Actions[0].Name);
    }

    [Fact]
    public void EnabledAndProjectDir_AreWorkbookScoped_SharedAcrossSheets()
    {
        var svc = new StateService();
        svc.SetCurrentSheet("wb", "Sheet1");
        svc.SetEnabled("wb", true);
        svc.SetProjectDir("wb", "/proj");

        svc.SetCurrentSheet("wb", "Sheet2");
        var s2 = svc.Get("wb");
        Assert.True(s2.Enabled);
        Assert.Equal("/proj", s2.ProjectDir);
    }

    [Fact]
    public void SettingAvailableScripts_DoesNotMaterialiseTheCurrentSheet()
    {
        // A workbook-scoped edit must not copy an inherited default into the
        // active sheet — only sheet-scoped edits create a sheet entry.
        var svc = new StateService();
        svc.SetCurrentSheet("wb", "Sheet1");
        svc.SetAvailableScripts("wb", new[] { "a.py" });

        var profile = svc.GetProfile("wb");
        Assert.False(profile.Sheets.ContainsKey("Sheet1"));
    }

    [Fact]
    public void GetProfile_SnapshotsWorkbookScopeAndAllConfiguredSheets()
    {
        var svc = new StateService();
        svc.SetEnabled("wb", true);
        svc.SetCurrentSheet("wb", "Sheet1");
        svc.SetPyInput("wb", "A1");
        svc.SetCurrentSheet("wb", "Sheet2");
        svc.SetPyOutput("wb", "B2");

        var p = svc.GetProfile("wb");
        Assert.True(p.Enabled);
        Assert.Equal("A1", p.Sheets["Sheet1"].PyInput);
        Assert.Equal("B2", p.Sheets["Sheet2"].PyOutput);
    }

    [Fact]
    public void LoadProfile_ReplacesStateAndProjectsTheCurrentSheet()
    {
        var svc = new StateService();
        var data = new WorkbookProfileData
        {
            Enabled = true,
            ProjectDir = "/proj",
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["S"] = new SheetProfile { SelectedScript = "x.py", PyInput = "A1" },
            },
        };
        svc.LoadProfile("wb", data);
        svc.SetCurrentSheet("wb", "S");

        var s = svc.Get("wb");
        Assert.True(s.Enabled);
        Assert.Equal("/proj", s.ProjectDir);
        Assert.Equal("x.py", s.SelectedScript);
        Assert.Equal("A1", s.PyInput);
    }

    [Fact]
    public void DefaultBucket_IsInheritedThenCopiedOnWrite()
    {
        var svc = new StateService();
        svc.LoadProfile("wb", new WorkbookProfileData
        {
            Enabled = true,
            Sheets = new Dictionary<string, SheetProfile>
            {
                [WorkbookProfileData.DefaultSheetKey] = new SheetProfile { SelectedScript = "def.py" },
            },
        });

        // A sheet with no entry of its own inherits the default.
        svc.SetCurrentSheet("wb", "AnySheet");
        Assert.Equal("def.py", svc.Get("wb").SelectedScript);

        // Editing one field copies the inherited profile forward (copy-on-write)
        // without disturbing the default that other sheets still inherit.
        svc.SetPyInput("wb", "A1");
        Assert.Equal("def.py", svc.Get("wb").SelectedScript);
        Assert.Equal("A1", svc.Get("wb").PyInput);

        var p = svc.GetProfile("wb");
        Assert.Equal("A1", p.Sheets["AnySheet"].PyInput);
        Assert.Equal("def.py", p.Sheets["AnySheet"].SelectedScript);
        Assert.Null(p.Sheets[WorkbookProfileData.DefaultSheetKey].PyInput);
    }

    [Fact]
    public void SetCurrentSheet_FiresOnlyWhenTheActiveSheetChanges()
    {
        var svc = new StateService();
        var count = 0;
        svc.StateChanged += (_, _) => count++;

        svc.SetCurrentSheet("wb", "S");   // "" -> "S": fires
        svc.SetCurrentSheet("wb", "S");   // unchanged: no fire
        Assert.Equal(1, count);
    }
}
