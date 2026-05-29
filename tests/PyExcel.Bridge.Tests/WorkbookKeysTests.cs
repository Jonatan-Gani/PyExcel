using System;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Tests for <see cref="WorkbookKeys"/> — the shared key-derivation that
/// <c>ExcelWorkbookContext</c> and the COM event sink both use. The whole
/// point of the class is that those two paths agree, so the rule must be
/// pinned: saved → FullName, unsaved → a session-scoped synthetic key.
/// </summary>
public class WorkbookKeysTests
{
    [Fact]
    public void Resolve_SavedWorkbook_UsesFullName()
    {
        var key = WorkbookKeys.Resolve("Book1.xlsx", @"C:\data", @"C:\data\Book1.xlsx");
        Assert.Equal(@"C:\data\Book1.xlsx", key);
    }

    [Fact]
    public void Resolve_UnsavedWorkbook_UsesSessionScopedName()
    {
        // Excel reports an empty Path for a new-but-unsaved workbook.
        var key = WorkbookKeys.Resolve("Book1", "", "Book1");
        Assert.Equal($"unsaved:{WorkbookKeys.SessionGuid}:Book1", key);
    }

    [Fact]
    public void Resolve_DistinctUnsavedNames_DoNotCollide()
    {
        var a = WorkbookKeys.Resolve("Book1", "", "Book1");
        var b = WorkbookKeys.Resolve("Book2", "", "Book2");
        Assert.NotEqual(a, b);
    }

    [Fact]
    public void Resolve_SameUnsavedName_IsStableWithinSession()
    {
        Assert.Equal(
            WorkbookKeys.Resolve("Book1", "", "Book1"),
            WorkbookKeys.Resolve("Book1", "", "Book1"));
    }

    [Fact]
    public void UnsavedKey_EmbedsTheSessionGuid()
    {
        Assert.Equal($"unsaved:{WorkbookKeys.SessionGuid}:Book7", WorkbookKeys.UnsavedKey("Book7"));
    }

    [Fact]
    public void UnsavedKey_NullName_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookKeys.UnsavedKey(null!));
    }

    [Fact]
    public void Resolve_NullName_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookKeys.Resolve(null!, "", ""));
    }

    [Fact]
    public void Resolve_NullFullName_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookKeys.Resolve("Book1", @"C:\x", null!));
    }
}
