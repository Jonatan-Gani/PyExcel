using System.Collections.Generic;
using System.Linq;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Round-trip and edge cases for <see cref="WorkbookProfileCodec"/> — the
/// per-sheet <see cref="WorkbookProfileData"/> serializer embedded (by
/// <see cref="ProjectProfileCodec"/>) in a workbook's CustomXMLPart.
/// </summary>
public class WorkbookProfileCodecTests
{
    private static WorkbookProfileData Sample() => new()
    {
        Enabled = true,
        ProjectDir = @"C:\proj",
        Sheets = new Dictionary<string, SheetProfile>
        {
            ["Sheet1"] = new SheetProfile
            {
                SelectedScript = "transform.py",
                PyInput = "A1:C10",
                PyOutput = "F1",
                SelectedActionName = "compute",
                Actions = new[]
                {
                    new RibbonAction("compute", "transform.py", "A1:C10", "F1",
                        new Dictionary<string, string> { ["factor"] = "5" }),
                },
                ImportInput = "data.csv",
            },
            ["Summary"] = new SheetProfile { PasteOutput = "B2" },
        },
    };

    [Fact]
    public void RoundTrip_PreservesWorkbookScopeAndEverySheet()
    {
        var xml = WorkbookProfileCodec.SerializeToString(Sample());
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.NotNull(back);
        Assert.True(back!.Enabled);
        Assert.Equal(@"C:\proj", back.ProjectDir);

        var s1 = back.Sheets["Sheet1"];
        Assert.Equal("transform.py", s1.SelectedScript);
        Assert.Equal("A1:C10", s1.PyInput);
        Assert.Equal("F1", s1.PyOutput);
        Assert.Equal("compute", s1.SelectedActionName);
        Assert.Equal("data.csv", s1.ImportInput);
        var a = Assert.Single(s1.Actions);
        Assert.Equal("compute", a.Name);
        Assert.Equal("5", a.Kwargs!["factor"]);

        Assert.Equal("B2", back.Sheets["Summary"].PasteOutput);
    }

    [Fact]
    public void RoundTrip_StructuredExportDefaults_PreservedPerSheet()
    {
        var data = new WorkbookProfileData
        {
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["Sheet1"] = new SheetProfile
                {
                    ExportInput = "A1:C10",
                    ExportFolder = @"C:\out",
                    ExportBaseName = "report",
                    ExportFormat = "tsv",
                    ExportTimestamp = "compact",
                },
            },
        };

        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));

        var s = back!.Sheets["Sheet1"];
        Assert.Equal("A1:C10", s.ExportInput);
        Assert.Equal(@"C:\out", s.ExportFolder);
        Assert.Equal("report", s.ExportBaseName);
        Assert.Equal("tsv", s.ExportFormat);
        Assert.Equal("compact", s.ExportTimestamp);
    }

    [Fact]
    public void Serialize_SheetWithOnlyExportBaseName_IsConfiguredAndKept()
    {
        // A new export-default field must count toward IsConfigured so a sheet
        // carrying only it is persisted rather than skipped as "empty".
        var data = new WorkbookProfileData
        {
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["Only"] = new SheetProfile { ExportBaseName = "weekly" },
            },
        };

        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.True(back!.Sheets.ContainsKey("Only"));
        Assert.Equal("weekly", back.Sheets["Only"].ExportBaseName);
    }

    [Fact]
    public void RoundTrip_DefaultBucketKey_Survives()
    {
        var data = new WorkbookProfileData
        {
            Sheets = new Dictionary<string, SheetProfile>
            {
                [WorkbookProfileData.DefaultSheetKey] = new SheetProfile { SelectedScript = "d.py" },
            },
        };
        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.Equal("d.py", back!.Sheets[WorkbookProfileData.DefaultSheetKey].SelectedScript);
    }

    [Fact]
    public void Serialize_SkipsUnconfiguredSheets()
    {
        var data = new WorkbookProfileData
        {
            Enabled = true,
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["Blank"] = SheetProfile.Empty,
                ["Real"] = new SheetProfile { PyInput = "A1" },
            },
        };
        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.False(back!.Sheets.ContainsKey("Blank"));
        Assert.True(back.Sheets.ContainsKey("Real"));
    }

    [Fact]
    public void Serialize_KwargsDeterministicRegardlessOfInsertionOrder()
    {
        WorkbookProfileData With(params (string k, string v)[] kw) => new()
        {
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["S"] = new SheetProfile
                {
                    Actions = new[]
                    {
                        new RibbonAction("a", "s.py", "A1", "B1",
                            kw.ToDictionary(p => p.k, p => p.v)),
                    },
                },
            },
        };

        Assert.Equal(
            WorkbookProfileCodec.SerializeToString(With(("c", "3"), ("a", "1"), ("b", "2"))),
            WorkbookProfileCodec.SerializeToString(With(("a", "1"), ("b", "2"), ("c", "3"))));
    }

    [Fact]
    public void RoundTrip_DefaultKeepOutputOpen_IsTrue()
    {
        // The Sample action doesn't set KeepOutputOpen, so it defaults to true
        // and must survive a round-trip as true.
        var xml = WorkbookProfileCodec.SerializeToString(Sample());
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.True(Assert.Single(back!.Sheets["Sheet1"].Actions).KeepOutputOpen);
    }

    [Fact]
    public void RoundTrip_KeepOutputOpenFalse_Survives()
    {
        var data = new WorkbookProfileData
        {
            Enabled = true,
            Sheets = new Dictionary<string, SheetProfile>
            {
                ["Sheet1"] = new SheetProfile
                {
                    Actions = new[]
                    {
                        new RibbonAction("a", "s.py", "A1", "B1",
                            Kwargs: null, KeepOutputOpen: false),
                    },
                },
            },
        };

        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.False(Assert.Single(back!.Sheets["Sheet1"].Actions).KeepOutputOpen);
    }

    [Fact]
    public void Deserialize_ActionWithoutKeepOutputOpenAttribute_DefaultsTrue()
    {
        // A document written before the attribute existed has no
        // keep-output-open on <action>; loading must default it to true.
        const string xml =
            "<workbook xmlns=\"urn:pyexcel:workbook:1\" version=\"1\">" +
            "<enabled>true</enabled>" +
            "<sheets><sheet name=\"Sheet1\"><actions>" +
            "<action name=\"a\" script=\"s.py\" input=\"A1\" output=\"B1\" />" +
            "</actions></sheet></sheets></workbook>";

        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.True(Assert.Single(back!.Sheets["Sheet1"].Actions).KeepOutputOpen);
    }

    [Fact]
    public void RoundTrip_Identity_ProjectIdAndOriginPathPreserved()
    {
        var data = new WorkbookProfileData
        {
            Enabled = true,
            ProjectId = "abc123def456",
            OriginPath = @"C:\proj\model.xlsx",
        };
        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.Equal("abc123def456", back!.ProjectId);
        Assert.Equal(@"C:\proj\model.xlsx", back.OriginPath);
    }

    [Fact]
    public void RoundTrip_NoIdentity_StaysNull()
    {
        var data = new WorkbookProfileData { Enabled = true };
        var xml = WorkbookProfileCodec.SerializeToString(data);
        Assert.True(WorkbookProfileCodec.TryDeserialize(xml, out var back));
        Assert.Null(back!.ProjectId);
        Assert.Null(back.OriginPath);
    }

    [Fact]
    public void TryDeserialize_Foreign_ReturnsFalse()
    {
        Assert.False(WorkbookProfileCodec.TryDeserialize("<nope/>", out var data));
        Assert.Null(data);
    }

    [Fact]
    public void TryDeserialize_NullOrBlank_ReturnsFalse()
    {
        Assert.False(WorkbookProfileCodec.TryDeserialize(null, out _));
        Assert.False(WorkbookProfileCodec.TryDeserialize("   ", out _));
    }
}
