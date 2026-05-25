namespace PyExcel.State;

/// <summary>
/// Abstraction over "which workbook is the user looking at right now."
/// In production this delegates to <c>Application.ActiveWorkbook</c> via
/// the COM API (see the Phase-3 follow-up <c>ExcelWorkbookContext</c>
/// implementation under <c>#if NETFRAMEWORK</c>). In tests, a fake
/// returns whatever string the test wants so ribbon-getter logic can
/// run cross-platform.
/// </summary>
public interface IWorkbookContext
{
    /// <summary>
    /// Stable key for the currently-active workbook, or
    /// <see langword="null"/> if Excel has no workbook in focus. The key
    /// must be the same string each time the user looks at the same
    /// workbook so <see cref="StateService"/> can identify it.
    /// </summary>
    /// <remarks>
    /// Production uses the full workbook path. New-but-unsaved workbooks
    /// have no path; the implementation falls back to the
    /// <c>Workbook.Name</c> plus a per-session GUID so two unsaved
    /// workbooks don't collide.
    /// </remarks>
    string? CurrentWorkbookKey { get; }
}

/// <summary>Null-object context — the ribbon then renders as if no
/// workbook were active. Used as the default before <see cref="PyExcel.Addin"/>
/// wires the real implementation.</summary>
public sealed class NullWorkbookContext : IWorkbookContext
{
    public static readonly NullWorkbookContext Instance = new();
    private NullWorkbookContext() { }
    public string? CurrentWorkbookKey => null;
}
