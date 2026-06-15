using System;
using System.Collections.Generic;
using System.IO;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Covers the cross-platform pieces of the workbook profile — the
/// <see cref="ProjectProfileCodec"/> round-trip (which nests the per-sheet
/// <see cref="WorkbookProfileCodec"/>) and the <see cref="ProjectMetadataFactory"/>
/// — that together make up the <see cref="ProjectProfile"/> embedded in a
/// workbook's <c>CustomXMLPart</c>. The COM read/write of the part itself is
/// Windows-only (<c>WorkbookStatePersister</c>) and exercised by the smoke test.
/// </summary>
public class ProjectProfileTests
{
    [Fact]
    public void Codec_RoundTrips_PerSheet_Profile_And_Metadata()
    {
        var data = new WorkbookProfileData
        {
            Enabled = true,
            ProjectDir = "/wb",
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["Sheet1"] = new SheetProfile
                {
                    SelectedScript = "s.py",
                    PyInput = "A1:C10",
                    Actions = new[] { new RibbonAction("a", "s.py", "A1", "B1") },
                    SelectedActionName = "a",
                },
                ["Data"] = new SheetProfile { ExportOutput = "out.csv" },
            },
        };
        var meta = new ProjectMetadata(PythonVersion: "3.12.1", WorkbookName: "Book.xlsx");

        var xml = ProjectProfileCodec.SerializeToString(data, meta);
        Assert.Contains("pyexcel-project", xml);

        Assert.True(ProjectProfileCodec.TryDeserialize(xml, "k", out var back, out var m));
        Assert.NotNull(back);
        Assert.True(back!.Enabled);
        Assert.Equal("/wb", back.ProjectDir);
        Assert.Equal("s.py", back.Sheets["Sheet1"].SelectedScript);
        Assert.Equal("A1:C10", back.Sheets["Sheet1"].PyInput);
        Assert.Single(back.Sheets["Sheet1"].Actions);
        Assert.Equal("a", back.Sheets["Sheet1"].SelectedActionName);
        Assert.Equal("out.csv", back.Sheets["Data"].ExportOutput);
        Assert.NotNull(m);
        Assert.Equal("3.12.1", m!.PythonVersion);
        Assert.Equal("Book.xlsx", m.WorkbookName);
    }

    [Fact]
    public void Codec_Migrates_Legacy_Flat_State_To_Default_Sheet()
    {
        // A project document an earlier build wrote: metadata + a nested flat
        // single-state <pyexcel> element (urn:pyexcel:state:1). It must migrate
        // forward into the default-bucket sheet so the workbook keeps its config.
        var flat = WorkbookState.Empty("k") with
        {
            Enabled = true,
            PyInput = "A1:B2",
            SelectedScript = "s.py",
        };
        var flatXml = WorkbookStateCodec.Serialize(flat).Root!.ToString();
        var projectXml =
            $"<pyexcel-project xmlns=\"{ProjectProfileCodec.XmlNamespace}\" project-version=\"1\">" +
            flatXml +
            "</pyexcel-project>";

        Assert.True(ProjectProfileCodec.TryDeserialize(projectXml, "k", out var data, out _));
        Assert.NotNull(data);
        Assert.True(data!.Enabled);
        var def = data.Sheets[WorkbookProfileData.DefaultSheetKey];
        Assert.Equal("A1:B2", def.PyInput);
        Assert.Equal("s.py", def.SelectedScript);
    }

    [Fact]
    public void Codec_Rejects_Foreign_Xml()
    {
        Assert.False(ProjectProfileCodec.TryDeserialize("<other/>", "k", out var data, out var meta));
        Assert.Null(data);
        Assert.Null(meta);
    }

    [Fact]
    public void MetadataFactory_Captures_Host_Info_And_Defaults_CreatedUtc()
    {
        var meta = ProjectMetadataFactory.Build(
            projectDir: null, workbookName: "Book.xlsx", workbookPath: "/wb/Book.xlsx", prior: null);

        Assert.False(string.IsNullOrEmpty(meta.GeneratedBy));
        Assert.NotNull(meta.Os);
        Assert.NotNull(meta.CreatedUtc);
        Assert.Equal(meta.CreatedUtc, meta.UpdatedUtc);
        Assert.Equal("Book.xlsx", meta.WorkbookName);
    }

    [Fact]
    public void MetadataFactory_Preserves_CreatedUtc_And_NonRecomputable_From_Prior()
    {
        var created = DateTimeOffset.UtcNow.AddDays(-3);
        var prior = new ProjectMetadata(CreatedUtc: created, WorkbookName: "Old.xlsx");

        var meta = ProjectMetadataFactory.Build(projectDir: null, workbookName: null, workbookPath: null, prior: prior);

        Assert.Equal(created, meta.CreatedUtc);
        Assert.True(meta.UpdatedUtc >= created);
        Assert.Equal("Old.xlsx", meta.WorkbookName);
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
