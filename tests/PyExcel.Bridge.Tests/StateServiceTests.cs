using System;
using System.Collections.Generic;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class StateServiceTests
{
    // -------------------------------------------------------------------------
    // Defaults
    // -------------------------------------------------------------------------

    [Fact]
    public void Get_UnknownWorkbook_ReturnsEmptyState()
    {
        var svc = new StateService();
        var s = svc.Get("never-seen");
        Assert.Equal("never-seen", s.WorkbookKey);
        Assert.False(s.Enabled);
        Assert.Empty(s.AvailableScripts);
        Assert.Empty(s.Actions);
        Assert.Null(s.SelectedScript);
        Assert.Null(s.SelectedActionName);
    }

    // -------------------------------------------------------------------------
    // Update + helpers
    // -------------------------------------------------------------------------

    [Fact]
    public void SetEnabled_UpdatesAndFiresStateChanged()
    {
        var svc = new StateService();
        string? seenKey = null;
        svc.StateChanged += (_, e) => seenKey = e.WorkbookKey;

        svc.SetEnabled("wb.xlsx", true);

        Assert.True(svc.Get("wb.xlsx").Enabled);
        Assert.Equal("wb.xlsx", seenKey);
    }

    [Fact]
    public void Mutator_PreservesUnchangedFields()
    {
        var svc = new StateService();
        svc.SetEnabled("wb.xlsx", true);
        svc.SetPyInput("wb.xlsx", "A1:C10");

        var s = svc.Get("wb.xlsx");
        Assert.True(s.Enabled);
        Assert.Equal("A1:C10", s.PyInput);
        Assert.Null(s.PyOutput);  // never set
    }

    [Fact]
    public void Mutator_ReturningNull_Throws()
    {
        var svc = new StateService();
        Assert.Throws<InvalidOperationException>(() =>
            svc.Update("wb.xlsx", _ => null!));
    }

    [Fact]
    public void Mutator_ChangingWorkbookKey_Throws()
    {
        var svc = new StateService();
        Assert.Throws<InvalidOperationException>(() =>
            svc.Update("wb.xlsx", s => s with { WorkbookKey = "evil.xlsx" }));
    }

    // -------------------------------------------------------------------------
    // Multiple workbooks isolation
    // -------------------------------------------------------------------------

    [Fact]
    public void Workbooks_AreIsolated()
    {
        var svc = new StateService();
        svc.SetEnabled("a.xlsx", true);
        svc.SetEnabled("b.xlsx", false);
        Assert.True(svc.Get("a.xlsx").Enabled);
        Assert.False(svc.Get("b.xlsx").Enabled);
    }

    [Fact]
    public void Forget_RemovesWorkbookAndFiresChanged()
    {
        var svc = new StateService();
        svc.SetEnabled("a.xlsx", true);
        Assert.Contains("a.xlsx", svc.KnownWorkbooks());

        string? seenKey = null;
        svc.StateChanged += (_, e) => seenKey = e.WorkbookKey;

        svc.Forget("a.xlsx");
        Assert.DoesNotContain("a.xlsx", svc.KnownWorkbooks());
        Assert.Equal("a.xlsx", seenKey);

        // Get after Forget returns Empty (not a stale cached version)
        Assert.False(svc.Get("a.xlsx").Enabled);
    }

    [Fact]
    public void Forget_UnknownWorkbook_DoesNotFireEvent()
    {
        var svc = new StateService();
        var seenCount = 0;
        svc.StateChanged += (_, _) => seenCount++;
        svc.Forget("never-seen");
        Assert.Equal(0, seenCount);
    }

    // -------------------------------------------------------------------------
    // Actions list management
    // -------------------------------------------------------------------------

    [Fact]
    public void AddAction_AppendsAndSelects()
    {
        var svc = new StateService();
        var action = new RibbonAction("first", "script.py", "A1", "B1");
        svc.AddAction("wb.xlsx", action);

        var s = svc.Get("wb.xlsx");
        Assert.Single(s.Actions);
        Assert.Equal("first", s.Actions[0].Name);
        Assert.Equal("first", s.SelectedActionName);
    }

    [Fact]
    public void LoadAction_CopiesScriptInputOutputIntoRunBoxesAndSelects()
    {
        var svc = new StateService();
        var action = new RibbonAction("calc", "calc.py", "Sheet1!A1:B2", "Sheet1!D1");

        svc.LoadAction("wb.xlsx", action);

        var s = svc.Get("wb.xlsx");
        Assert.Equal("calc.py", s.SelectedScript);
        Assert.Equal("Sheet1!A1:B2", s.PyInput);
        Assert.Equal("Sheet1!D1", s.PyOutput);
        Assert.Equal("calc", s.SelectedActionName);
    }

    [Fact]
    public void AddAction_SameName_Upserts()
    {
        var svc = new StateService();
        svc.AddAction("wb.xlsx", new RibbonAction("a", "v1.py", "A1", "B1"));
        svc.AddAction("wb.xlsx", new RibbonAction("a", "v2.py", "C1", "D1"));

        var s = svc.Get("wb.xlsx");
        Assert.Single(s.Actions);
        Assert.Equal("v2.py", s.Actions[0].Script);
    }

    [Fact]
    public void DeleteAction_RemovesAndClearsSelectionIfMatched()
    {
        var svc = new StateService();
        svc.AddAction("wb.xlsx", new RibbonAction("a", "a.py", "A1", "B1"));
        svc.AddAction("wb.xlsx", new RibbonAction("b", "b.py", "C1", "D1"));
        Assert.Equal("b", svc.Get("wb.xlsx").SelectedActionName);

        svc.DeleteAction("wb.xlsx", "b");
        var s = svc.Get("wb.xlsx");
        Assert.Single(s.Actions);
        Assert.Equal("a", s.Actions[0].Name);
        Assert.Null(s.SelectedActionName);
    }

    [Fact]
    public void DeleteAction_OtherSelected_KeepsSelection()
    {
        var svc = new StateService();
        svc.AddAction("wb.xlsx", new RibbonAction("a", "a.py", "A1", "B1"));
        svc.AddAction("wb.xlsx", new RibbonAction("b", "b.py", "C1", "D1"));
        svc.SetSelectedAction("wb.xlsx", "a");

        svc.DeleteAction("wb.xlsx", "b");
        Assert.Equal("a", svc.Get("wb.xlsx").SelectedActionName);
    }

    [Fact]
    public void SelectedAction_ResolvesToActualRecord()
    {
        var svc = new StateService();
        svc.AddAction("wb.xlsx", new RibbonAction("a", "a.py", "A1", "B1"));
        svc.AddAction("wb.xlsx", new RibbonAction("b", "b.py", "C1", "D1"));
        svc.SetSelectedAction("wb.xlsx", "a");

        var selected = svc.Get("wb.xlsx").SelectedAction;
        Assert.NotNull(selected);
        Assert.Equal("a.py", selected!.Script);
    }

    [Fact]
    public void SelectedAction_NameDoesNotResolve_ReturnsNull()
    {
        var svc = new StateService();
        svc.SetSelectedAction("wb.xlsx", "ghost");
        Assert.Null(svc.Get("wb.xlsx").SelectedAction);
    }

    // -------------------------------------------------------------------------
    // Script list management
    // -------------------------------------------------------------------------

    [Fact]
    public void SetAvailableScripts_ReplacesEntireList()
    {
        var svc = new StateService();
        svc.SetAvailableScripts("wb.xlsx", new[] { "a", "b" });
        svc.SetAvailableScripts("wb.xlsx", new[] { "x", "y", "z" });
        var s = svc.Get("wb.xlsx");
        Assert.Equal(new[] { "x", "y", "z" }, s.AvailableScripts);
    }

    // -------------------------------------------------------------------------
    // Argument validation
    // -------------------------------------------------------------------------

    [Fact]
    public void Get_NullKey_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new StateService().Get(null!));
    }

    [Fact]
    public void Update_NullKey_Throws()
    {
        Assert.Throws<ArgumentNullException>(() =>
            new StateService().Update(null!, s => s));
    }

    [Fact]
    public void Update_NullMutator_Throws()
    {
        Assert.Throws<ArgumentNullException>(() =>
            new StateService().Update("wb.xlsx", null!));
    }

    // -------------------------------------------------------------------------
    // Phase 5 — Import / Export / Paste text fields
    // -------------------------------------------------------------------------

    [Fact]
    public void SetImportInput_PersistsAndPreservesOtherFields()
    {
        var svc = new StateService();
        svc.SetEnabled("wb.xlsx", true);
        svc.SetPyInput("wb.xlsx", "A1");

        svc.SetImportInput("wb.xlsx", "data.csv");

        var s = svc.Get("wb.xlsx");
        Assert.Equal("data.csv", s.ImportInput);
        Assert.True(s.Enabled);
        Assert.Equal("A1", s.PyInput);
    }

    [Fact]
    public void SetImportOutput_PersistsValue()
    {
        var svc = new StateService();
        svc.SetImportOutput("wb.xlsx", "Sheet1!A1");
        Assert.Equal("Sheet1!A1", svc.Get("wb.xlsx").ImportOutput);
    }

    [Fact]
    public void SetExportInput_PersistsValue()
    {
        var svc = new StateService();
        svc.SetExportInput("wb.xlsx", "A1:C10");
        Assert.Equal("A1:C10", svc.Get("wb.xlsx").ExportInput);
    }

    [Fact]
    public void SetExportOutput_PersistsValue()
    {
        var svc = new StateService();
        svc.SetExportOutput("wb.xlsx", "out.csv");
        Assert.Equal("out.csv", svc.Get("wb.xlsx").ExportOutput);
    }

    [Fact]
    public void SetPasteOutput_PersistsValue()
    {
        var svc = new StateService();
        svc.SetPasteOutput("wb.xlsx", "Sheet1!B2");
        Assert.Equal("Sheet1!B2", svc.Get("wb.xlsx").PasteOutput);
    }

    [Fact]
    public void SetImportInput_FiresStateChanged()
    {
        var svc = new StateService();
        string? seenKey = null;
        svc.StateChanged += (_, e) => seenKey = e.WorkbookKey;

        svc.SetImportInput("wb.xlsx", "data.csv");
        Assert.Equal("wb.xlsx", seenKey);
    }

    [Fact]
    public void SetImportInput_Null_ClearsField()
    {
        var svc = new StateService();
        svc.SetImportInput("wb.xlsx", "data.csv");
        svc.SetImportInput("wb.xlsx", null);
        Assert.Null(svc.Get("wb.xlsx").ImportInput);
    }

    [Fact]
    public void NewSetters_NullKey_Throws()
    {
        var svc = new StateService();
        Assert.Throws<ArgumentNullException>(() => svc.SetImportInput(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => svc.SetImportOutput(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => svc.SetExportInput(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => svc.SetExportOutput(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => svc.SetPasteOutput(null!, "x"));
    }
}
