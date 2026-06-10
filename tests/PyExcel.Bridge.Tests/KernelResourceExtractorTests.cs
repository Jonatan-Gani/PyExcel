using System;
using System.IO;
using System.Linq;
using PyExcel.Setup.Kernel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Tests for <see cref="KernelResourceExtractor"/> using the actual
/// PyExcel.Setup assembly as the resource source — the same assembly
/// the .xll ships with, so a regression in the csproj's embed list
/// surfaces here before any user sees it.
/// </summary>
public class KernelResourceExtractorTests
{
    [Fact]
    public void EnumerateKernelResources_ShipsCanonicalSet()
    {
        var extractor = new KernelResourceExtractor();
        var names = extractor.EnumerateKernelResources().ToList();

        // Every kernel module the v2 bridge currently relies on must be
        // embedded; if a file is added to embedded/pyexcel/kernel/ a
        // matching <EmbeddedResource> entry must land in the csproj.
        Assert.Contains("pyexcel/__init__.py", names);
        Assert.Contains("pyexcel/kernel/__init__.py", names);
        Assert.Contains("pyexcel/kernel/__main__.py", names);
        Assert.Contains("pyexcel/kernel/arrow_io.py", names);
        Assert.Contains("pyexcel/kernel/chart.py", names);
        Assert.Contains("pyexcel/kernel/framing.py", names);
        Assert.Contains("pyexcel/kernel/supervisor.py", names);
        Assert.Contains("pyexcel/kernel/transport.py", names);
        Assert.Contains("pyexcel/kernel/types.py", names);
        Assert.Contains("pyexcel/kernel/worker.py", names);
        Assert.Contains("pyexcel/requirements.txt", names);
    }

    [Fact]
    public void Extract_WritesFilesToDisk()
    {
        var target = NewTempDir();
        try
        {
            var result = new KernelResourceExtractor().Extract(target);

            Assert.Equal(target, result.TargetDir);
            Assert.NotEmpty(result.Written);
            Assert.Empty(result.Skipped);

            // The __main__ entrypoint is the canonical kernel marker
            // PyExcel.Excel.PythonResolver.ResolveEmbeddedPath looks
            // for — the layout must put it at
            // `<target>/pyexcel/kernel/__main__.py` so adding <target>
            // to PYTHONPATH makes `import pyexcel.kernel` resolve.
            var mainPy = Path.Combine(target, "pyexcel", "kernel", "__main__.py");
            Assert.True(File.Exists(mainPy));
            var bytes = File.ReadAllBytes(mainPy);
            Assert.NotEmpty(bytes);

            // The package marker must exist too — the parent __init__.py
            // is what makes `pyexcel` a package on disk.
            var pkgInit = Path.Combine(target, "pyexcel", "__init__.py");
            Assert.True(File.Exists(pkgInit));

            // requirements.txt rides as a sibling resource inside the
            // package directory.
            var req = Path.Combine(target, "pyexcel", "requirements.txt");
            Assert.True(File.Exists(req));
        }
        finally
        {
            Cleanup(target);
        }
    }

    [Fact]
    public void Extract_SecondCall_SkipsUnchangedFiles()
    {
        var target = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(target);

            var second = extractor.Extract(target);
            Assert.Empty(second.Written);
            Assert.NotEmpty(second.Skipped);
        }
        finally
        {
            Cleanup(target);
        }
    }

    [Fact]
    public void Extract_OverwritesTamperedFile()
    {
        var target = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(target);

            var mainPy = Path.Combine(target, "pyexcel", "kernel", "__main__.py");
            File.WriteAllText(mainPy, "# tampered");

            var second = extractor.Extract(target);
            // The tampered file must appear in Written — content differs
            // from the embedded copy, so the extractor refreshes it.
            var expectedRelative = string.Join(Path.DirectorySeparatorChar.ToString(),
                new[] { "pyexcel", "kernel", "__main__.py" });
            Assert.Contains(expectedRelative, second.Written);
            Assert.DoesNotContain("# tampered", File.ReadAllText(mainPy));
        }
        finally
        {
            Cleanup(target);
        }
    }

    [Fact]
    public void Extract_NullOrEmptyTarget_Throws()
    {
        var extractor = new KernelResourceExtractor();
        Assert.Throws<ArgumentException>(() => extractor.Extract(string.Empty));
        Assert.Throws<ArgumentException>(() => extractor.Extract("  "));
    }

    private static string NewTempDir()
    {
        var dir = Path.Combine(
            Path.GetTempPath(),
            "pyexcel-kernel-extract-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static void Cleanup(string dir)
    {
        try { Directory.Delete(dir, recursive: true); }
        catch { /* best-effort */ }
    }
}
