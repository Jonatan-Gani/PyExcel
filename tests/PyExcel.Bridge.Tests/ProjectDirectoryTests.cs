using System;
using System.IO;
using PyExcel.Common;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ProjectDirectoryTests
{
    // Run each case with a known-clean override, restoring whatever was there.
    private static void WithOverride(string? value, Action body)
    {
        var original = Environment.GetEnvironmentVariable(ProjectDirectory.OverrideEnvVar);
        try
        {
            Environment.SetEnvironmentVariable(ProjectDirectory.OverrideEnvVar, value);
            body();
        }
        finally
        {
            Environment.SetEnvironmentVariable(ProjectDirectory.OverrideEnvVar, original);
        }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("https://contoso.sharepoint.com/sites/Team/Shared Documents")]
    [InlineData("http://server/share/book")]
    public void IsUsableLocalPath_RejectsBlankAndUrls(string? path)
    {
        Assert.False(ProjectDirectory.IsUsableLocalPath(path));
    }

    [Fact]
    public void IsUsableLocalPath_AcceptsARealLocalDirectory()
    {
        Assert.True(ProjectDirectory.IsUsableLocalPath(Path.GetTempPath()));
    }

    [Fact]
    public void Resolve_LocalWorkbookDir_ReturnsItNormalised()
    {
        WithOverride(null, () =>
        {
            var dir = Path.GetTempPath();
            Assert.Equal(Path.GetFullPath(dir), ProjectDirectory.Resolve(dir));
        });
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Resolve_BlankWithoutOverride_PassesThrough(string? dir)
    {
        WithOverride(null, () => Assert.Equal(dir, ProjectDirectory.Resolve(dir)));
    }

    [Fact]
    public void Resolve_OverrideWins_EvenForLocalAndUrl()
    {
        var pin = Path.GetTempPath();
        WithOverride(pin, () =>
        {
            var expected = Path.GetFullPath(pin);
            Assert.Equal(expected, ProjectDirectory.Resolve(@"https://contoso.sharepoint.com/x"));
            Assert.Equal(expected, ProjectDirectory.Resolve(Path.GetTempPath()));
            Assert.Equal(expected, ProjectDirectory.Resolve(null));
        });
    }

    [Fact]
    public void Resolve_SharePointUrl_FallsBackToLocalDirNotTheUrl()
    {
        WithOverride(null, () =>
        {
            const string url = "https://contoso.sharepoint.com/sites/Team/Shared Documents/Reports";
            var resolved = ProjectDirectory.Resolve(url);

            Assert.NotNull(resolved);
            Assert.NotEqual(url, resolved);
            // The fallback is itself a usable local directory and is namespaced.
            Assert.True(ProjectDirectory.IsUsableLocalPath(resolved));
            Assert.Contains("PyExcel", resolved!);
        });
    }

    [Fact]
    public void Resolve_SharePointUrl_IsDeterministicAndDistinct()
    {
        WithOverride(null, () =>
        {
            const string a = "https://contoso.sharepoint.com/sites/A/Docs";
            const string b = "https://contoso.sharepoint.com/sites/B/Docs";

            Assert.Equal(ProjectDirectory.Resolve(a), ProjectDirectory.Resolve(a));
            Assert.NotEqual(ProjectDirectory.Resolve(a), ProjectDirectory.Resolve(b));
        });
    }

    // -------------------------------------------------------------------------
    // Stored-project-dir overload (self-healing location)
    // -------------------------------------------------------------------------

    [Fact]
    public void ResolveStored_Existing_ReturnsStored()
    {
        WithOverride(null, () =>
        {
            var stored = Path.Combine(Path.GetTempPath(), "proj");
            var wb = Path.Combine(Path.GetTempPath(), "wbfolder");
            Assert.Equal(
                Path.GetFullPath(stored),
                ProjectDirectory.Resolve(stored, wb, directoryExists: _ => true));
        });
    }

    [Fact]
    public void ResolveStored_Stale_FallsBackToWorkbookFolder()
    {
        // The stored project folder no longer exists (moved wholesale) — resolve to
        // the workbook's own folder, which moved with it.
        WithOverride(null, () =>
        {
            var stored = Path.Combine(Path.GetTempPath(), "moved-away");
            var wb = Path.GetTempPath();
            Assert.Equal(
                Path.GetFullPath(wb),
                ProjectDirectory.Resolve(stored, wb, directoryExists: _ => false));
        });
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void ResolveStored_None_UsesWorkbookFolder(string? stored)
    {
        WithOverride(null, () =>
        {
            var wb = Path.GetTempPath();
            Assert.Equal(
                Path.GetFullPath(wb),
                ProjectDirectory.Resolve(stored, wb, directoryExists: _ => true));
        });
    }

    [Fact]
    public void ResolveStored_Url_FallsBackToWorkbookFolder()
    {
        WithOverride(null, () =>
        {
            var wb = Path.GetTempPath();
            Assert.Equal(
                Path.GetFullPath(wb),
                ProjectDirectory.Resolve(
                    "https://contoso.sharepoint.com/x", wb, directoryExists: _ => true));
        });
    }

    [Fact]
    public void ResolveStored_OverrideWins_OverStored()
    {
        var pin = Path.GetTempPath();
        WithOverride(pin, () =>
        {
            // Even though stored 'exists', the global override pins the location.
            var stored = Path.Combine(Path.GetTempPath(), "proj");
            Assert.Equal(
                Path.GetFullPath(pin),
                ProjectDirectory.Resolve(stored, "/some/wb", directoryExists: _ => true));
        });
    }
}
