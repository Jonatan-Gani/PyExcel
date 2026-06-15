using System;
using System.IO;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers the cross-platform pieces of the workbook profile — the
/// <see cref="ProjectProfileCodec"/> round-trip and the
/// <see cref="ProjectMetadataFactory"/> — that together make up the
/// <see cref="ProjectProfile"/> embedded in a workbook's <c>CustomXMLPart</c>
/// (the single source of truth for "is this workbook a PyExcel project?"). The
/// COM read/write of the part itself is Windows-only (<c>WorkbookStatePersister</c>)
/// and exercised by the Windows smoke test.
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
        // The locator namespace must be present or the COM persister could never
        // find the part again via SelectByNamespace.
        Assert.Contains(ProjectProfileCodec.XmlNamespace, xml);

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
    public void Codec_KeysByTheCallerSuppliedKey()
    {
        // A workbook saved-as-a-copy gets a new key; the embedded XML never
        // carried the key, so the caller's key must win on load.
        var xml = ProjectProfileCodec.SerializeToString(
            WorkbookState.Empty("old-key") with { Enabled = true }, new ProjectMetadata());
        Assert.True(ProjectProfileCodec.TryDeserialize(xml, "new-key", out var s, out _));
        Assert.Equal("new-key", s!.WorkbookKey);
    }

    [Fact]
    public void Codec_Rejects_Foreign_Xml()
    {
        Assert.False(ProjectProfileCodec.TryDeserialize("<other/>", "k", out var s, out var m));
        Assert.Null(s);
        Assert.Null(m);
    }

    [Fact]
    public void MetadataFactory_Captures_Host_Info_And_Defaults_CreatedUtc()
    {
        var meta = ProjectMetadataFactory.Build(
            projectDir: null, workbookName: "Book.xlsx", workbookPath: "/wb/Book.xlsx", prior: null);

        Assert.False(string.IsNullOrEmpty(meta.GeneratedBy));
        Assert.NotNull(meta.Os);
        Assert.NotNull(meta.CreatedUtc);
        // First build: created and updated are stamped from the same instant.
        Assert.Equal(meta.CreatedUtc, meta.UpdatedUtc);
        Assert.Equal("Book.xlsx", meta.WorkbookName);
    }

    [Fact]
    public void MetadataFactory_Preserves_CreatedUtc_And_NonRecomputable_From_Prior()
    {
        var created = DateTimeOffset.UtcNow.AddDays(-3);
        var prior = new ProjectMetadata(CreatedUtc: created, WorkbookName: "Old.xlsx");

        var meta = ProjectMetadataFactory.Build(projectDir: null, workbookName: null, workbookPath: null, prior: prior);

        Assert.Equal(created, meta.CreatedUtc);       // preserved across saves
        Assert.True(meta.UpdatedUtc >= created);      // refreshed each save
        Assert.Equal("Old.xlsx", meta.WorkbookName);  // kept when not supplied fresh
    }

    [Fact]
    public void MetadataFactory_Reads_Venv_Python_From_PyvenvCfg()
    {
        var dir = NewTempDir();
        try
        {
            var venv = Path.Combine(dir, ".pyexcel-venv");
            Directory.CreateDirectory(Path.Combine(venv, "bin"));
            File.WriteAllText(Path.Combine(venv, "pyvenv.cfg"), "home = /usr/bin\nversion = 3.12.1\n");
            File.WriteAllText(Path.Combine(venv, "bin", "python"), "#!stub");

            var meta = ProjectMetadataFactory.Build(dir, "Book.xlsx", workbookPath: null, prior: null);

            Assert.Equal("3.12.1", meta.PythonVersion);
            Assert.NotNull(meta.PythonPath);
            Assert.EndsWith("python", meta.PythonPath!);
        }
        finally { Directory.Delete(dir, true); }
    }

    private static string NewTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pyexcel-profile-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return dir;
    }
}
