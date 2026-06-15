using System;
using System.IO;
using System.Linq;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Fast, file-only validation that an enabled workbook's project structure
/// (venv, kernel, userScripts) is present — what the open hook checks so the user
/// is told the environment is missing before they hit Run.
/// </summary>
public class ProjectStructureValidatorTests
{
    [Fact]
    public void Validate_AllPresent_IsOk()
    {
        var dir = NewProject(venv: true, kernel: true, scripts: true);
        try
        {
            var r = ProjectStructureValidator.Validate(dir);
            Assert.True(r.Ok);
            Assert.Empty(r.Missing);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Validate_WindowsVenvLayout_IsRecognised()
    {
        var dir = NewProject(venv: false, kernel: true, scripts: true);
        try
        {
            // Windows layout: Scripts\python.exe instead of bin/python.
            var scripts = Path.Combine(dir, ProjectStructureValidator.VenvDirName, "Scripts");
            Directory.CreateDirectory(scripts);
            File.WriteAllText(Path.Combine(scripts, "python.exe"), "stub");

            var r = ProjectStructureValidator.Validate(dir);
            Assert.True(r.Ok);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Validate_MissingVenv_ReportsOnlyThat()
    {
        var dir = NewProject(venv: false, kernel: true, scripts: true);
        try
        {
            var r = ProjectStructureValidator.Validate(dir);
            Assert.False(r.Ok);
            Assert.Single(r.Missing);
            Assert.Contains(".pyexcel-venv", r.Missing.Single());
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Validate_EmptyProject_ReportsEveryComponent()
    {
        var dir = NewProject(venv: false, kernel: false, scripts: false);
        try
        {
            var r = ProjectStructureValidator.Validate(dir);
            Assert.False(r.Ok);
            Assert.Equal(3, r.Missing.Count);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Validate_NoProjectDir_IsNotOk(string? dir)
    {
        var r = ProjectStructureValidator.Validate(dir);
        Assert.False(r.Ok);
        Assert.Contains("project folder", r.Missing.Single());
    }

    [Fact]
    public void Validate_NonexistentDir_IsNotOk()
    {
        var r = ProjectStructureValidator.Validate(Path.Combine(Path.GetTempPath(), "pyexcel-nope-" + Guid.NewGuid().ToString("N")));
        Assert.False(r.Ok);
    }

    private static string NewProject(bool venv, bool kernel, bool scripts)
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-struct-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        if (venv)
        {
            var bin = Path.Combine(dir, ProjectStructureValidator.VenvDirName, "bin");
            Directory.CreateDirectory(bin);
            File.WriteAllText(Path.Combine(bin, "python"), "stub");
        }
        if (kernel)
        {
            var k = Path.Combine(dir, ProjectStructureValidator.KernelDirName, "pyexcel", "kernel");
            Directory.CreateDirectory(k);
            File.WriteAllText(Path.Combine(k, "__main__.py"), "stub");
        }
        if (scripts)
            Directory.CreateDirectory(Path.Combine(dir, ProjectStructureValidator.ScriptsDirName));
        return dir;
    }
}
