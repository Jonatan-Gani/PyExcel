using System;
using System.IO;
using PyExcel.Excel;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class ImportPlannerTests
{
    // -------------------------------------------------------------------------
    // Field validation
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankImportInput_Throws(string? input)
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create(input, "A1", workbookDirectory: null));
        Assert.Contains("Input", ex.Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Create_BlankImportOutput_Throws(string? output)
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create("data.csv", output, workbookDirectory: null));
        Assert.Contains("Output", ex.Message);
    }

    [Fact]
    public void Create_TrimsLeadingAndTrailingWhitespace()
    {
        var plan = ImportPlanner.Create(
            "  data.csv  ",
            "  Sheet1!A1  ",
            workbookDirectory: "/wb");
        Assert.Equal("Sheet1!A1", plan.TargetRangeAddress);
        // Whitespace stripped before the path was joined to /wb.
        Assert.EndsWith("data.csv", plan.AbsoluteSourcePath);
    }

    // -------------------------------------------------------------------------
    // Path resolution
    // -------------------------------------------------------------------------

    [Fact]
    public void ResolvePath_AbsoluteIsPreserved()
    {
        var abs = OperatingSystem_IsWindows() ? @"C:\data\foo.csv" : "/data/foo.csv";
        var resolved = ImportPlanner.ResolvePath(abs, workbookDirectory: "/elsewhere");
        Assert.Equal(Path.GetFullPath(abs), resolved);
    }

    [Fact]
    public void ResolvePath_RelativeJoinsToWorkbookDir()
    {
        var basis = OperatingSystem_IsWindows() ? @"C:\wb" : "/wb";
        var resolved = ImportPlanner.ResolvePath("data.csv", workbookDirectory: basis);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(basis, "data.csv")),
            resolved);
    }

    [Fact]
    public void ResolvePath_RelativeNoWorkbookDir_UsesCurrentDirectory()
    {
        var resolved = ImportPlanner.ResolvePath("data.csv", workbookDirectory: null);
        Assert.Equal(
            Path.GetFullPath(Path.Combine(Environment.CurrentDirectory, "data.csv")),
            resolved);
    }

    [Fact]
    public void ResolvePath_BlankSource_Throws()
    {
        Assert.Throws<ArgumentException>(() => ImportPlanner.ResolvePath("", "/wb"));
    }

    // -------------------------------------------------------------------------
    // Delimiter detection
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("foo.csv", ',')]
    [InlineData("foo.txt", ',')]
    [InlineData("foo", ',')]
    [InlineData("foo.CSV", ',')]
    public void DetectDelimiter_DefaultsToComma(string path, char expected)
    {
        Assert.Equal(expected, ImportPlanner.DetectDelimiter(path));
    }

    [Theory]
    [InlineData("foo.tsv")]
    [InlineData("foo.TSV")]
    [InlineData("/path/to/foo.tsv")]
    public void DetectDelimiter_TsvExtension_ReturnsTab(string path)
    {
        Assert.Equal('\t', ImportPlanner.DetectDelimiter(path));
    }

    [Theory]
    [InlineData("foo.xlsx")]
    [InlineData("foo.xls")]
    [InlineData("foo.xlsm")]
    [InlineData("foo.xlsb")]
    [InlineData("foo.ods")]
    [InlineData("foo.XLSX")]
    public void DetectDelimiter_BinaryFormats_Throws(string path)
    {
        // DetectDelimiter is reused by ExportPlanner, which still rejects
        // Excel-format targets. Excel-format imports route through
        // DetectFormat instead — the planner's Create method never calls
        // DetectDelimiter for an Excel path.
        Assert.Throws<FormatException>(() => ImportPlanner.DetectDelimiter(path));
    }

    [Fact]
    public void DetectDelimiter_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ImportPlanner.DetectDelimiter(null!));
    }

    // -------------------------------------------------------------------------
    // Format detection
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData("foo.csv")]
    [InlineData("foo.tsv")]
    [InlineData("foo.txt")]
    [InlineData("foo")]
    [InlineData("foo.CSV")]
    public void DetectFormat_TextLikeExtensions_ReturnsCsv(string path)
    {
        Assert.Equal(ImportFormat.Csv, ImportPlanner.DetectFormat(path));
    }

    [Theory]
    [InlineData("foo.xlsx")]
    [InlineData("foo.xlsm")]
    [InlineData("foo.xlsb")]
    [InlineData("foo.XLSX")]
    [InlineData("/path/to/foo.xlsm")]
    public void DetectFormat_ModernExcelExtensions_ReturnsExcel(string path)
    {
        Assert.Equal(ImportFormat.Excel, ImportPlanner.DetectFormat(path));
    }

    [Fact]
    public void DetectFormat_XlsLegacyBinary_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.DetectFormat("foo.xls"));
        Assert.Contains(".xls", ex.Message);
    }

    [Fact]
    public void DetectFormat_OdsOpenDocument_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.DetectFormat("foo.ods"));
        Assert.Contains(".ods", ex.Message);
    }

    [Fact]
    public void DetectFormat_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ImportPlanner.DetectFormat(null!));
    }

    // -------------------------------------------------------------------------
    // path!Sheet syntax parsing
    // -------------------------------------------------------------------------

    [Fact]
    public void ParsePathAndSheet_NoBang_ReturnsWhole()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx");
        Assert.Equal("data.xlsx", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_ExcelExtensionWithSheet_Splits()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx!Q2");
        Assert.Equal("data.xlsx", path);
        Assert.Equal("Q2", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_XlsmWithSheet_Splits()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("model.xlsm!Inputs");
        Assert.Equal("model.xlsm", path);
        Assert.Equal("Inputs", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_XlsbWithSheet_Splits()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("big.xlsb!Data");
        Assert.Equal("big.xlsb", path);
        Assert.Equal("Data", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_SheetSeparatorOnly_NoSheet()
    {
        // 'data.xlsx!' — the user typed the separator with no sheet name.
        // Treated as "default sheet" (null), not as a sheet named "".
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx!");
        Assert.Equal("data.xlsx", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_WhitespaceOnlySheet_TreatedAsNoSheet()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx!   ");
        Assert.Equal("data.xlsx", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_TrimsSheetWhitespace()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx!  Inputs  ");
        Assert.Equal("data.xlsx", path);
        Assert.Equal("Inputs", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_CsvWithBang_DoesNotSplit()
    {
        // .csv isn't an Excel format — '!' in the name stays in the path.
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("weird!name.csv");
        Assert.Equal("weird!name.csv", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_CsvSheetSyntax_DoesNotSplit()
    {
        // CSV doesn't have sheets — the planner's Create() will reject
        // this with a clean FormatException, but the parser itself just
        // hands the whole thing back as the path.
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.csv!Sheet1");
        Assert.Equal("data.csv!Sheet1", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_NoExtensionWithBang_DoesNotSplit()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data!Sheet1");
        Assert.Equal("data!Sheet1", path);
        Assert.Null(sheet);
    }

    [Fact]
    public void ParsePathAndSheet_PathWithDirectoryAndSheet_Splits()
    {
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("subdir/data.xlsx!Sheet1");
        Assert.Equal("subdir/data.xlsx", path);
        Assert.Equal("Sheet1", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_SheetNameWithSpaces_PreservedInside()
    {
        // 'My Inputs' — internal spaces are part of the sheet name and stay.
        var (path, sheet) = ImportPlanner.ParsePathAndSheet("data.xlsx!My Inputs");
        Assert.Equal("data.xlsx", path);
        Assert.Equal("My Inputs", sheet);
    }

    [Fact]
    public void ParsePathAndSheet_NullInput_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ImportPlanner.ParsePathAndSheet(null!));
    }

    // -------------------------------------------------------------------------
    // Compose — inverse of ParsePathAndSheet (drives the sheet picker's
    // write-back into the Import field)
    // -------------------------------------------------------------------------

    [Fact]
    public void Compose_NullPath_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => ImportPlanner.Compose(null!, "Q2"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Compose_BlankSheet_ReturnsPathUnchanged(string? sheet)
    {
        Assert.Equal("data.xlsx", ImportPlanner.Compose("data.xlsx", sheet));
    }

    [Fact]
    public void Compose_ExcelPathWithSheet_AppendsBangSheet()
    {
        Assert.Equal("data.xlsx!Q2", ImportPlanner.Compose("data.xlsx", "Q2"));
    }

    [Fact]
    public void Compose_TrimsSheet()
    {
        Assert.Equal("data.xlsx!Q2", ImportPlanner.Compose("data.xlsx", "  Q2  "));
    }

    [Theory]
    [InlineData("data.csv")]
    [InlineData("data.tsv")]
    [InlineData("data.txt")]
    [InlineData("data")]
    public void Compose_SheetOnNonExcelPath_Throws(string path)
    {
        Assert.Throws<FormatException>(() => ImportPlanner.Compose(path, "Q2"));
    }

    [Theory]
    [InlineData("data.xlsx", "Q2")]
    [InlineData("model.xlsm", "Inputs")]
    [InlineData("big.xlsb", "Data")]
    [InlineData("data.xlsx", "My Inputs")]
    public void Compose_RoundTripsWithParsePathAndSheet(string path, string sheet)
    {
        var composed = ImportPlanner.Compose(path, sheet);
        var (parsedPath, parsedSheet) = ImportPlanner.ParsePathAndSheet(composed);
        Assert.Equal(path, parsedPath);
        Assert.Equal(sheet, parsedSheet);
    }

    // -------------------------------------------------------------------------
    // End-to-end Create() composition
    // -------------------------------------------------------------------------

    [Fact]
    public void Create_TsvExtension_PicksTabDelimiter()
    {
        var plan = ImportPlanner.Create("data.tsv", "A1", workbookDirectory: "/wb");
        Assert.Equal('\t', plan.Delimiter);
        Assert.Equal(ImportFormat.Csv, plan.Format);
        Assert.Null(plan.SheetName);
    }

    [Fact]
    public void Create_CsvExtension_PicksCommaDelimiter()
    {
        var plan = ImportPlanner.Create("data.csv", "A1", workbookDirectory: "/wb");
        Assert.Equal(',', plan.Delimiter);
        Assert.Equal(ImportFormat.Csv, plan.Format);
        Assert.Null(plan.SheetName);
    }

    [Fact]
    public void Create_XlsxExtension_RoutesToExcelFormat()
    {
        var plan = ImportPlanner.Create("data.xlsx", "A1", workbookDirectory: "/wb");
        Assert.Equal(ImportFormat.Excel, plan.Format);
        Assert.Null(plan.SheetName);
        Assert.EndsWith("data.xlsx", plan.AbsoluteSourcePath);
    }

    [Fact]
    public void Create_XlsxWithSheetSyntax_RoutesToExcelWithSheetName()
    {
        var plan = ImportPlanner.Create("data.xlsx!Inputs", "A1", workbookDirectory: "/wb");
        Assert.Equal(ImportFormat.Excel, plan.Format);
        Assert.Equal("Inputs", plan.SheetName);
        Assert.EndsWith("data.xlsx", plan.AbsoluteSourcePath);
        Assert.DoesNotContain("!", plan.AbsoluteSourcePath);
    }

    [Fact]
    public void Create_XlsmWithSheetSyntax_RoutesToExcelWithSheetName()
    {
        var plan = ImportPlanner.Create("model.xlsm!Q2", "A1", workbookDirectory: "/wb");
        Assert.Equal(ImportFormat.Excel, plan.Format);
        Assert.Equal("Q2", plan.SheetName);
    }

    [Fact]
    public void Create_XlsbWithSheetSyntax_RoutesToExcelWithSheetName()
    {
        var plan = ImportPlanner.Create("big.xlsb!Data", "A1", workbookDirectory: "/wb");
        Assert.Equal(ImportFormat.Excel, plan.Format);
        Assert.Equal("Data", plan.SheetName);
    }

    [Fact]
    public void Create_XlsExtension_Throws()
    {
        // Legacy .xls binary is explicitly rejected — the COM importer
        // only supports the modern Excel binaries.
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create("legacy.xls", "A1", workbookDirectory: "/wb"));
        Assert.Contains(".xls", ex.Message);
    }

    [Fact]
    public void Create_OdsExtension_Throws()
    {
        var ex = Assert.Throws<FormatException>(
            () => ImportPlanner.Create("doc.ods", "A1", workbookDirectory: "/wb"));
        Assert.Contains(".ods", ex.Message);
    }

    [Fact]
    public void Create_CsvPathContainingBang_IsPreservedNotSplit()
    {
        // ParsePathAndSheet only splits on '!' when the part before is
        // an Excel extension, so 'data.csv!Sheet1' stays whole — the
        // planner resolves it as a (likely non-existent) CSV path. The
        // file-not-found error at read time is the right surface for
        // this; the planner shouldn't pre-fail on the syntax.
        var plan = ImportPlanner.Create("data.csv!Sheet1", "A1", workbookDirectory: "/wb");
        Assert.Equal(ImportFormat.Csv, plan.Format);
        Assert.Null(plan.SheetName);
        Assert.EndsWith("data.csv!Sheet1", plan.AbsoluteSourcePath);
    }

    /// <summary>Tiny helper — `OperatingSystem.IsWindows()` lives in
    /// net5+; the test project targets net8 so it's available, but
    /// rolling our own keeps the test source independent of which TFM
    /// the test assembly happens to be on.</summary>
    private static bool OperatingSystem_IsWindows()
        => Environment.OSVersion.Platform is PlatformID.Win32NT
            or PlatformID.Win32S
            or PlatformID.Win32Windows
            or PlatformID.WinCE;
}
