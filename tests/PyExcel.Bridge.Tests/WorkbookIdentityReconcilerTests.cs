using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class WorkbookIdentityReconcilerTests
{
    [Fact]
    public void Unstamped_IsUnchanged()
        => Assert.Equal(
            WorkbookIdentityAction.Unchanged,
            WorkbookIdentityReconciler.Reconcile(
                projectId: null, originPath: "/a/wb.xlsx", currentPath: "/b/wb.xlsx", originExists: true));

    [Fact]
    public void NoCommittedOrigin_IsUnchanged()
        => Assert.Equal(
            WorkbookIdentityAction.Unchanged,
            WorkbookIdentityReconciler.Reconcile("g", originPath: null, currentPath: "/b/wb.xlsx", originExists: false));

    [Fact]
    public void NoCurrentPath_IsUnchanged()
        => Assert.Equal(
            WorkbookIdentityAction.Unchanged,
            WorkbookIdentityReconciler.Reconcile("g", "/a/wb.xlsx", currentPath: null, originExists: true));

    [Fact]
    public void SamePath_IsUnchanged()
        => Assert.Equal(
            WorkbookIdentityAction.Unchanged,
            WorkbookIdentityReconciler.Reconcile("g", "/a/wb.xlsx", "/a/wb.xlsx", originExists: true));

    [Fact]
    public void PathChanged_OriginGone_IsMoved()
        => Assert.Equal(
            WorkbookIdentityAction.Moved,
            WorkbookIdentityReconciler.Reconcile("g", "/a/wb.xlsx", "/b/wb.xlsx", originExists: false));

    [Fact]
    public void PathChanged_OriginStillExists_IsCopied()
        => Assert.Equal(
            WorkbookIdentityAction.Copied,
            WorkbookIdentityReconciler.Reconcile("g", "/a/wb.xlsx", "/b/wb.xlsx", originExists: true));

    [Fact]
    public void PathsEqual_NormalisesRelativeSegments()
        => Assert.True(WorkbookIdentityReconciler.PathsEqual("/a/b/../wb.xlsx", "/a/wb.xlsx"));

    [Fact]
    public void PathsEqual_IsCaseInsensitive()
        => Assert.True(WorkbookIdentityReconciler.PathsEqual("/A/WB.xlsx", "/a/wb.xlsx"));

    [Fact]
    public void PathsEqual_DifferentPaths_NotEqual()
        => Assert.False(WorkbookIdentityReconciler.PathsEqual("/a/wb.xlsx", "/b/wb.xlsx"));
}
