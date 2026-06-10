using System;
using System.IO;
using System.Runtime.InteropServices;
using PyExcel.Setup.Paths;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="ProjectPathResolver"/> tests. The classifier is pure
/// (no disk reads), so every case is a string transformation.
/// </summary>
public class ProjectPathResolverTests
{
    [Fact]
    public void Resolve_NullOrWhitespace_Throws()
    {
        var resolver = new ProjectPathResolver();
        Assert.Throws<ArgumentException>(() => resolver.Resolve(string.Empty));
        Assert.Throws<ArgumentException>(() => resolver.Resolve("   "));
    }

    [Fact]
    public void Resolve_PlainLocalPath_NormalisedAndNotFlagged()
    {
        var resolver = new ProjectPathResolver();
        var temp = Path.GetTempPath();
        var info = resolver.Resolve(temp);

        Assert.False(info.IsUnc);
        Assert.False(info.IsOneDriveSynced);
        Assert.Null(info.OneDriveRoot);
        Assert.Equal(temp, info.OriginalPath);
        // Normalised path drops trailing separators and resolves
        // any . / .. segments; we just assert it has the same prefix.
        Assert.StartsWith(temp.TrimEnd(Path.DirectorySeparatorChar),
            info.NormalisedPath.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Resolve_OneDrivePath_FlagsAsSynced()
    {
        // Drive the resolver against a fake OneDrive root so the test
        // works regardless of the host's actual OneDrive state.
        var fakeRoot = Path.Combine(Path.GetTempPath(), "FakeOneDrive_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(fakeRoot);
        var prior = Environment.GetEnvironmentVariable("OneDrive");
        try
        {
            Environment.SetEnvironmentVariable("OneDrive", fakeRoot);
            var nested = Path.Combine(fakeRoot, "Projects", "PyExcel");

            var info = new ProjectPathResolver().Resolve(nested);
            Assert.True(info.IsOneDriveSynced);
            Assert.NotNull(info.OneDriveRoot);
            Assert.StartsWith(
                Path.GetFullPath(fakeRoot).TrimEnd(Path.DirectorySeparatorChar),
                info.OneDriveRoot!.TrimEnd(Path.DirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OneDrive", prior);
            try { Directory.Delete(fakeRoot, recursive: true); } catch { /* swallow */ }
        }
    }

    [Fact]
    public void Resolve_PathOutsideOneDrive_NotFlaggedEvenWhenOneDriveSet()
    {
        var fakeRoot = Path.Combine(Path.GetTempPath(), "FakeOneDrive_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(fakeRoot);
        var prior = Environment.GetEnvironmentVariable("OneDrive");
        try
        {
            Environment.SetEnvironmentVariable("OneDrive", fakeRoot);
            var elsewhere = Path.Combine(Path.GetTempPath(), "outside-" + Guid.NewGuid().ToString("N"));
            var info = new ProjectPathResolver().Resolve(elsewhere);
            Assert.False(info.IsOneDriveSynced);
            Assert.Null(info.OneDriveRoot);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OneDrive", prior);
            try { Directory.Delete(fakeRoot, recursive: true); } catch { /* swallow */ }
        }
    }

    [Fact]
    public void Resolve_PathThatLooksLikeOneDrivePrefix_NotFlagged()
    {
        // If OneDrive root is `<tmp>/SyncRoot_xyz`, then a sibling
        // directory `<tmp>/SyncRoot_xyzExtra` shares the string prefix
        // but is NOT a child of the OneDrive root. The resolver checks
        // for a directory boundary, not a raw string prefix, so the
        // sibling must not be flagged.
        var stem = "SyncRoot_" + Guid.NewGuid().ToString("N");
        var fakeRoot = Path.Combine(Path.GetTempPath(), stem);
        var siblingRoot = Path.Combine(Path.GetTempPath(), stem + "Extra");
        Directory.CreateDirectory(fakeRoot);
        Directory.CreateDirectory(siblingRoot);
        var prior = Environment.GetEnvironmentVariable("OneDrive");
        try
        {
            Environment.SetEnvironmentVariable("OneDrive", fakeRoot);
            var info = new ProjectPathResolver().Resolve(siblingRoot);
            Assert.False(info.IsOneDriveSynced);
        }
        finally
        {
            Environment.SetEnvironmentVariable("OneDrive", prior);
            try { Directory.Delete(fakeRoot, recursive: true); } catch { /* swallow */ }
            try { Directory.Delete(siblingRoot, recursive: true); } catch { /* swallow */ }
        }
    }

    [Fact]
    public void Resolve_UncPath_FlaggedOnWindows()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return; // UNC is a Windows-only concept.

        // We cannot stat a nonexistent UNC path during normalisation,
        // but Path.GetFullPath does not actually hit the network for
        // a well-formed `\\server\share\…` form on Windows. If the
        // particular CI image rejects it for some reason, the test
        // surfaces it as an ArgumentException and we update accordingly.
        var info = new ProjectPathResolver().Resolve(@"\\server\share\project");
        Assert.True(info.IsUnc);
    }

    [Fact]
    public void Resolve_LongUncPrefix_IsNormalisedToCanonicalForm()
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return;

        var info = new ProjectPathResolver().Resolve(@"\\?\UNC\server\share\project");
        Assert.True(info.IsUnc);
        Assert.StartsWith(@"\\server\share", info.NormalisedPath, StringComparison.OrdinalIgnoreCase);
    }
}
