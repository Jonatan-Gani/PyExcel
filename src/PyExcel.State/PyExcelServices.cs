namespace PyExcel.State;

/// <summary>
/// Service locator for the ribbon. Excel-DNA instantiates
/// <c>PyExcelRibbon</c> via a parameterless constructor, so the ribbon
/// can't take its dependencies via DI — it pulls them off a static
/// container instead.
///
/// <para>The add-in's <c>AutoOpen</c> wires the real services here
/// before the first ribbon callback can fire. Tests replace them with
/// fakes (using a try/finally to restore the originals).</para>
///
/// <para>The defaults are safe to use unconfigured: an empty
/// <see cref="StateService"/> and a <see cref="NullWorkbookContext"/>
/// produce a ribbon that renders as "no workbook, all disabled" —
/// which is exactly what we want until <c>AutoOpen</c> runs.</para>
/// </summary>
public static class PyExcelServices
{
    /// <summary>The process-wide state registry.</summary>
    public static StateService State { get; set; } = new StateService();

    /// <summary>Strategy for "what workbook is active right now".</summary>
    public static IWorkbookContext WorkbookContext { get; set; } = NullWorkbookContext.Instance;
}
