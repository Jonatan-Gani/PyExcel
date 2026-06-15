using System;
using System.IO;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers the project-folder profile — <see cref="ProjectProfileCodec"/> and
/// <see cref="ProjectProfileStore"/> — the single authoritative, portable record
/// that makes an enabled workbook come back enabled on reopen.
/// </summary>
public class ProjectProfileTests
{
    [Fact]
    public void Codec_RoundTrips_State_And_Metadata()
    {
        var key = "/wb/Book.xlsx";
        var state = WorkbookState.Empty(key) with
        {
            Enabled = true,
            ProjectDir = "/wb",
            Actions = new[] { new RibbonAction("a", "s.py", "in", "out") },
        };
        var meta = new ProjectMetadata(
            GeneratedBy: "PyExcel 2.0.0.0",
            PythonVersion: "3.12.1",
            WorkbookName: "Book.xlsx");

        var xml = ProjectProfileCodec.SerializeToString(state, meta);
        Assert.Contains("pyexcel-project", xml);

        Assert.True(ProjectProfileCodec.TryDeserialize(xml, key, out var s, out var m));
        Assert.NotNull(s);
        Assert.True(s!.Enabled);
        Assert.Equal("/wb", s.ProjectDir);
        Assert.Single(s.Actions);
        Assert.Equal("a", s.Actions[0].Name);
        Assert.NotNull(m);
        Assert.Equal("3.12.1", m!.PythonVersion);
        Assert.Equal("Book.xlsx", m.WorkbookName);
    }

    [Fact]
    public void Codec_Rejects_Foreign_Xml()
    {
        Assert.False(ProjectProfileCodec.TryDeserialize("<other/>", "k", out var s, out var m));
        Assert.Null(s);
        Assert.Null(m);
    }

    [Fact]
    public void Store_Save_Then_Load_From_Project_Folder()
    {
        var dir = NewTempDir();
        try
        {
            var key = Path.Combine(dir, "Book.xlsx");
            var state = WorkbookState.Empty(key) with { Enabled = true, ProjectDir = dir };

            ProjectProfileStore.Save(dir, state, "Book.xlsx", key);

            // The profile lands inside the .pyexcel subfolder, not as a loose
            // .xml next to the workbook (which Excel would try to open).
            var file = ProjectProfileStore.PathFor(dir);
            Assert.NotNull(file);
            Assert.True(File.Exists(file!));
            Assert.Equal(
                Path.Combine(dir, ProjectProfileStore.SubDirName, ProjectProfileStore.FileName), file);
            Assert.False(File.Exists(Path.Combine(dir, "pyexcel.project.xml")));

            var loaded = ProjectProfileStore.TryLoad(dir, key);
            Assert.NotNull(loaded);
            Assert.True(loaded!.Enabled);
            Assert.Equal(dir, loaded.ProjectDir);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Store_Captures_Host_Metadata_And_Preserves_CreatedUtc()
    {
        var dir = NewTempDir();
        try
        {
            var key = Path.Combine(dir, "Book.xlsx");
            var state = WorkbookState.Empty(key) with { Enabled = true, ProjectDir = dir };

            ProjectProfileStore.Save(dir, state, "Book.xlsx", key);
            var first = ProjectProfileStore.TryLoadProfile(dir, key);
            Assert.NotNull(first);
            Assert.False(string.IsNullOrEmpty(first!.Metadata.GeneratedBy));
            Assert.NotNull(first.Metadata.CreatedUtc);
            Assert.NotNull(first.Metadata.Os);

            // A second save keeps the original created-utc but refreshes updated-utc.
            ProjectProfileStore.Save(dir, state, "Book.xlsx", key);
            var second = ProjectProfileStore.TryLoadProfile(dir, key);
            Assert.Equal(first.Metadata.CreatedUtc, second!.Metadata.CreatedUtc);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Store_Reads_Legacy_Root_File_And_Migrates_On_Save()
    {
        var dir = NewTempDir();
        try
        {
            var key = Path.Combine(dir, "Book.xlsx");
            var state = WorkbookState.Empty(key) with { Enabled = true, ProjectDir = dir };

            // An earlier build's loose profile next to the workbook.
            var legacy = Path.Combine(dir, "pyexcel.project.xml");
            File.WriteAllText(legacy, ProjectProfileCodec.SerializeToString(state, new ProjectMetadata()));

            // Readable via the fallback so existing projects keep working...
            var loaded = ProjectProfileStore.TryLoad(dir, key);
            Assert.NotNull(loaded);
            Assert.True(loaded!.Enabled);

            // ...and the next save migrates it into the subfolder and deletes the loose file.
            ProjectProfileStore.Save(dir, state, "Book.xlsx", key);
            Assert.False(File.Exists(legacy));
            Assert.True(File.Exists(ProjectProfileStore.PathFor(dir)!));
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Store_TryLoad_Missing_Returns_Null()
    {
        var dir = NewTempDir();
        try { Assert.Null(ProjectProfileStore.TryLoad(dir, "/x/Book.xlsx")); }
        finally { Directory.Delete(dir, true); }
    }

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-profile-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}
