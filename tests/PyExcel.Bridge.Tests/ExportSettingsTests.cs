using System.IO;
using PyExcel.Excel;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ExportSettingsTests
{
    // -------------------------------------------------------------------------
    // Token round-trips
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(ExportFileType.Csv, "csv")]
    [InlineData(ExportFileType.Tsv, "tsv")]
    public void FileTypeToken_RoundTrips(ExportFileType type, string token)
    {
        Assert.Equal(token, ExportSettings.ToToken(type));
        Assert.Equal(type, ExportSettings.ParseFileType(token));
    }

    [Theory]
    [InlineData(ExportTimestampStyle.None, "none")]
    [InlineData(ExportTimestampStyle.DateAndTime, "datetime")]
    [InlineData(ExportTimestampStyle.DateOnly, "date")]
    [InlineData(ExportTimestampStyle.Compact, "compact")]
    public void TimestampToken_RoundTrips(ExportTimestampStyle style, string token)
    {
        Assert.Equal(token, ExportSettings.ToToken(style));
        Assert.Equal(style, ExportSettings.ParseTimestamp(token));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("xlsx")]
    public void ParseFileType_Unknown_DefaultsCsv(string? token)
        => Assert.Equal(ExportFileType.Csv, ExportSettings.ParseFileType(token));

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("weekly")]
    public void ParseTimestamp_Unknown_DefaultsNone(string? token)
        => Assert.Equal(ExportTimestampStyle.None, ExportSettings.ParseTimestamp(token));

    [Fact]
    public void ParseFileType_IsCaseInsensitive()
        => Assert.Equal(ExportFileType.Tsv, ExportSettings.ParseFileType("TSV"));

    // -------------------------------------------------------------------------
    // FromState
    // -------------------------------------------------------------------------

    [Fact]
    public void FromState_StructuredFields_AreRead()
    {
        var state = WorkbookState.Empty("wb") with
        {
            ExportInput = "A1:C10",
            ExportFolder = @"C:\out",
            ExportBaseName = "report",
            ExportFormat = "tsv",
            ExportTimestamp = "datetime",
        };

        var s = ExportSettings.FromState(state);

        Assert.Equal("A1:C10", s.SourceRange);
        Assert.Equal(@"C:\out", s.Folder);
        Assert.Equal("report", s.BaseName);
        Assert.Equal(ExportFileType.Tsv, s.FileType);
        Assert.Equal(ExportTimestampStyle.DateAndTime, s.Timestamp);
    }

    [Fact]
    public void FromState_Empty_IsAllDefaults()
    {
        var s = ExportSettings.FromState(WorkbookState.Empty("wb"));
        Assert.Null(s.SourceRange);
        Assert.Null(s.Folder);
        Assert.Null(s.BaseName);
        Assert.Equal(ExportFileType.Csv, s.FileType);
        Assert.Equal(ExportTimestampStyle.None, s.Timestamp);
    }

    [Fact]
    public void FromState_LegacyExportOutput_DecomposedIntoFolderNameAndType()
    {
        // A workbook configured before the structured defaults existed only has
        // the single ExportOutput path — split it so nothing is lost on upgrade.
        const string legacyPath = "/data/exports/results.tsv";
        var state = WorkbookState.Empty("wb") with
        {
            ExportInput = "A1",
            ExportOutput = legacyPath,
        };

        var s = ExportSettings.FromState(state);

        Assert.Equal("A1", s.SourceRange);
        // GetDirectoryName normalises to the platform separator (back-slashes on
        // Windows), so assert against it rather than a hard-coded forward-slash form.
        Assert.Equal(Path.GetDirectoryName(legacyPath), s.Folder);
        Assert.Equal("results", s.BaseName);
        Assert.Equal(ExportFileType.Tsv, s.FileType);
        Assert.Equal(ExportTimestampStyle.None, s.Timestamp);
    }

    [Fact]
    public void FromState_StructuredFields_WinOverLegacyExportOutput()
    {
        var state = WorkbookState.Empty("wb") with
        {
            ExportBaseName = "structured",
            ExportFormat = "csv",
            ExportOutput = "/old/legacy.tsv",
        };

        var s = ExportSettings.FromState(state);

        Assert.Equal("structured", s.BaseName);
        Assert.Equal(ExportFileType.Csv, s.FileType);
        // The legacy folder is not consulted once a structured field is present.
        Assert.Null(s.Folder);
    }
}
