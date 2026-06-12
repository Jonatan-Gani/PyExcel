using System;
using System.IO;
using PyExcel.Forms;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers <see cref="ScriptScaffold"/> — the cross-platform half of the Note 2
/// "New script" button: name sanitising, the starter template, and
/// collision-safe creation on disk.
/// </summary>
public class ScriptScaffoldTests
{
    [Theory]
    [InlineData("myscript", "myscript.py")]
    [InlineData("my script", "my_script.py")]
    [InlineData("report.py", "report.py")]
    [InlineData("Report.PY", "Report.py")]
    [InlineData("a/b\\c", "a_b_c.py")]
    [InlineData("  trim  ", "trim.py")]
    public void SanitizeFileName_ProducesSafeName(string input, string expected)
        => Assert.Equal(expected, ScriptScaffold.SanitizeFileName(input));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("***")]
    public void SanitizeFileName_RejectsUnusable(string? input)
        => Assert.Throws<ArgumentException>(() => ScriptScaffold.SanitizeFileName(input));

    [Fact]
    public void Create_WritesTemplate_AndReturnsFileName()
    {
        var dir = NewTempDir();
        try
        {
            var name = ScriptScaffold.Create(dir, "demo");
            Assert.Equal("demo.py", name);
            var path = Path.Combine(dir, name);
            Assert.True(File.Exists(path));
            var text = File.ReadAllText(path);
            Assert.Contains("def transform(inputs:", text);
            Assert.Equal(ScriptScaffold.Template, text);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Create_CreatesMissingDirectory()
    {
        var parent = NewTempDir();
        try
        {
            var dir = Path.Combine(parent, "userScripts");
            var name = ScriptScaffold.Create(dir, "x");
            Assert.True(File.Exists(Path.Combine(dir, name)));
        }
        finally { Directory.Delete(parent, true); }
    }

    [Fact]
    public void Create_DisambiguatesCollisions()
    {
        var dir = NewTempDir();
        try
        {
            Assert.Equal("dup.py", ScriptScaffold.Create(dir, "dup"));
            Assert.Equal("dup_1.py", ScriptScaffold.Create(dir, "dup"));
            Assert.Equal("dup_2.py", ScriptScaffold.Create(dir, "dup"));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_RejectsEmptyDir(string dir)
        => Assert.Throws<ArgumentException>(() => ScriptScaffold.Create(dir, "x"));

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}
