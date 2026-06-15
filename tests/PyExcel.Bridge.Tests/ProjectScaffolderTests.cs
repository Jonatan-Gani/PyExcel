using System;
using System.IO;
using PyExcel.Setup;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers <see cref="ProjectScaffolder"/> — the Setup stage that lays out the
/// user-facing <c>userScripts</c> folder and a starter script next to the
/// workbook, the piece that was missing when Setup "built nothing around the
/// workbook".
/// </summary>
public class ProjectScaffolderTests
{
    [Fact]
    public void Scaffold_CreatesUserScriptsAndExample()
    {
        var dir = NewTempDir();
        try
        {
            var userScripts = new ProjectScaffolder().Scaffold(dir);
            Assert.Equal(Path.Combine(dir, "userScripts"), userScripts);
            Assert.True(Directory.Exists(userScripts));
            var example = Path.Combine(userScripts, "example.py");
            Assert.True(File.Exists(example));
            Assert.Contains("def transform(inputs:", File.ReadAllText(example));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Scaffold_DoesNotOverwriteExistingScripts()
    {
        var dir = NewTempDir();
        try
        {
            var userScripts = Path.Combine(dir, "userScripts");
            Directory.CreateDirectory(userScripts);
            var mine = Path.Combine(userScripts, "mine.py");
            File.WriteAllText(mine, "# mine");

            new ProjectScaffolder().Scaffold(dir);

            // No example dropped (the folder already has a script), and the
            // existing one is untouched.
            Assert.False(File.Exists(Path.Combine(userScripts, "example.py")));
            Assert.Equal("# mine", File.ReadAllText(mine));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Scaffold_IsIdempotent_KeepsEditedExample()
    {
        var dir = NewTempDir();
        try
        {
            var scaffolder = new ProjectScaffolder();
            var userScripts = scaffolder.Scaffold(dir);
            var example = Path.Combine(userScripts, "example.py");
            File.WriteAllText(example, "# edited");

            scaffolder.Scaffold(dir); // second run must not clobber

            Assert.Equal("# edited", File.ReadAllText(example));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Scaffold_WritesReadme()
    {
        var dir = NewTempDir();
        try
        {
            new ProjectScaffolder().Scaffold(dir);
            var readme = Path.Combine(dir, "README.md");
            Assert.True(File.Exists(readme), "Scaffold should write a README the Read Me button can open");
            Assert.Contains("PyExcel", File.ReadAllText(readme));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Scaffold_DoesNotOverwriteExistingReadme()
    {
        var dir = NewTempDir();
        try
        {
            var readme = Path.Combine(dir, "README.md");
            File.WriteAllText(readme, "# mine");
            new ProjectScaffolder().Scaffold(dir);
            Assert.Equal("# mine", File.ReadAllText(readme));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    public void Scaffold_RejectsBlankDir(string dir)
        => Assert.Throws<ArgumentException>(() => new ProjectScaffolder().Scaffold(dir));

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-scaffold-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}
