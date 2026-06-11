using System;
using System.IO;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// <see cref="PythonResolver.ResolveEmbeddedPath(string?)"/> tests, focused
/// on the Phase-7 wiring: a Setup-extracted <c>.pyexcel-kernel</c> under the
/// workbook directory is preferred over the bundled <c>embedded/</c>, and the
/// resolver falls back to the bundled copy when no extraction is present.
/// </summary>
public class PythonResolverTests
{
    [Fact]
    public void ResolveEmbeddedPath_NoWorkbookDir_FallsBackToBundledEmbedded()
    {
        // With no workbook directory the resolver walks up to the repo-root
        // embedded/ that the test host ships with — the same path production
        // uses for the copy bundled beside the .xll.
        var resolved = PythonResolver.ResolveEmbeddedPath();

        Assert.EndsWith("embedded", resolved.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine(resolved, "pyexcel", "kernel", "__main__.py")));
    }

    [Fact]
    public void ResolveEmbeddedPath_ExtractedKernelPresent_PrefersItOverBundled()
    {
        using var temp = new TempDir();
        var extracted = Path.Combine(temp.Path, PythonResolver.ExtractedKernelDirName);
        WriteKernelMarker(extracted);

        var resolved = PythonResolver.ResolveEmbeddedPath(temp.Path);

        Assert.Equal(extracted, resolved);
    }

    [Fact]
    public void ResolveEmbeddedPath_WorkbookDirWithoutExtraction_FallsBackToBundled()
    {
        using var temp = new TempDir();
        // No .pyexcel-kernel under the workbook dir → bundled copy wins.
        var resolved = PythonResolver.ResolveEmbeddedPath(temp.Path);

        Assert.EndsWith("embedded", resolved.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.Ordinal);
        Assert.NotEqual(
            Path.Combine(temp.Path, PythonResolver.ExtractedKernelDirName), resolved);
    }

    [Fact]
    public void ResolveEmbeddedPath_PartialExtraction_DoesNotMatchAndFallsBack()
    {
        using var temp = new TempDir();
        // A .pyexcel-kernel directory exists but the kernel marker is
        // absent (e.g. an interrupted extraction). It must not be picked.
        Directory.CreateDirectory(
            Path.Combine(temp.Path, PythonResolver.ExtractedKernelDirName, "pyexcel"));

        var resolved = PythonResolver.ResolveEmbeddedPath(temp.Path);

        Assert.NotEqual(
            Path.Combine(temp.Path, PythonResolver.ExtractedKernelDirName), resolved);
        Assert.EndsWith("embedded", resolved.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ResolveEmbeddedPath_BlankWorkbookDir_IgnoresItAndFallsBack(string? workbookDir)
    {
        var resolved = PythonResolver.ResolveEmbeddedPath(workbookDir);

        Assert.EndsWith("embedded", resolved.TrimEnd(Path.DirectorySeparatorChar),
            StringComparison.Ordinal);
    }

    private static void WriteKernelMarker(string root)
    {
        var dir = Path.Combine(root, "pyexcel", "kernel");
        Directory.CreateDirectory(dir);
        File.WriteAllText(Path.Combine(dir, "__main__.py"), "# test marker\n");
    }

    private sealed class TempDir : IDisposable
    {
        public string Path { get; }

        public TempDir()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "pyexcel-resolvertest-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public void Dispose()
        {
            try { Directory.Delete(Path, recursive: true); }
            catch { /* best-effort cleanup */ }
        }
    }
}
