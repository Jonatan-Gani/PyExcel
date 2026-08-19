using System.IO;
using System.Linq;
using PyExcel.Setup.Kernel;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// The staleness check behind the Update button.
///
/// <para>The gap these cover: the kernel that runs is the copy extracted into
/// the project folder, and nothing re-extracts it on its own. A user who
/// installs a new build keeps running the old kernel — it handshakes fine and
/// ignores meta it does not understand — so the two halves diverge silently.
/// Detecting that is what makes Update offerable.</para>
/// </summary>
public class KernelFreshnessTests
{
    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-fresh-" + Path.GetRandomFileName());
        Directory.CreateDirectory(dir);
        return dir;
    }

    [Fact]
    public void Check_ReportsUpToDate_ImmediatelyAfterExtraction()
    {
        var dir = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(dir);

            var check = extractor.Check(dir);

            Assert.True(check.UpToDate, check.Describe());
            Assert.Empty(check.Stale);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Check_NamesAFileWhoseContentDrifted()
    {
        var dir = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(dir);

            // Simulate the real failure: an older kernel left on disk while the
            // add-in moved on. Byte content differs; the version string need not.
            var worker = Path.Combine(dir, "pyexcel", "kernel", "worker.py");
            File.WriteAllText(worker, "# an older worker\n");

            var check = extractor.Check(dir);

            Assert.False(check.UpToDate);
            Assert.Contains(check.Stale, p => p.EndsWith("worker.py"));
            Assert.Contains("worker.py", check.Describe());
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Check_ReportsAMissingFile()
    {
        var dir = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(dir);

            var target = Path.Combine(dir, "pyexcel", "kernel", "declared_types.py");
            Assert.True(File.Exists(target), "declared_types.py should have been extracted");
            File.Delete(target);

            var check = extractor.Check(dir);

            Assert.False(check.UpToDate);
            Assert.Contains(check.Stale, p => p.EndsWith("declared_types.py"));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Check_TreatsAnEmptyDirectoryAsEntirelyStale()
    {
        var dir = NewTempDir();
        try
        {
            var check = new KernelResourceExtractor().Check(dir);

            Assert.False(check.UpToDate);
            Assert.Equal(
                new KernelResourceExtractor().EnumerateKernelResources().Count(),
                check.Stale.Count);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void ExtractingAgain_MakesAStaleKernelFreshWithoutTouchingTheRest()
    {
        var dir = NewTempDir();
        try
        {
            var extractor = new KernelResourceExtractor();
            extractor.Extract(dir);
            File.WriteAllText(Path.Combine(dir, "pyexcel", "kernel", "worker.py"), "# stale\n");

            var result = extractor.Extract(dir);

            Assert.True(extractor.Check(dir).UpToDate);
            // Only the drifted file is rewritten — this is what makes Update
            // safe to press on a working project.
            Assert.Contains(result.Written, p => p.EndsWith("worker.py"));
            Assert.True(result.Skipped.Count > 0);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void IsUpToDate_IsTrueWhenNoKernelHasBeenExtractedYet()
    {
        var dir = NewTempDir();
        try
        {
            // Nothing to update — that is Enable's job, not Update's.
            KernelFreshness.Invalidate(dir);
            Assert.True(KernelFreshness.IsUpToDate(dir));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void IsUpToDate_DetectsDriftAndClearsAfterInvalidate()
    {
        var dir = NewTempDir();
        try
        {
            var kernelDir = Path.Combine(dir, KernelFreshness.KernelDirName);
            var extractor = new KernelResourceExtractor();
            extractor.Extract(kernelDir);

            KernelFreshness.Invalidate(dir);
            Assert.True(KernelFreshness.IsUpToDate(dir));

            File.WriteAllText(
                Path.Combine(kernelDir, "pyexcel", "kernel", "worker.py"), "# stale\n");

            // Still cached from the previous probe — the memo is the whole point.
            Assert.True(KernelFreshness.IsUpToDate(dir));

            KernelFreshness.Invalidate(dir);
            Assert.False(KernelFreshness.IsUpToDate(dir));
        }
        finally
        {
            KernelFreshness.Invalidate(dir);
            Directory.Delete(dir, true);
        }
    }
}
