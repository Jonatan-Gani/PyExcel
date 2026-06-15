using System;
using System.Collections.Concurrent;
using System.IO;

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

    /// <summary>The process-wide last-error registry. Wired by the
    /// add-in's <c>AutoOpen</c>; the ribbon reads from it to render
    /// the Show / Copy Last Error buttons.</summary>
    public static ErrorService Errors { get; set; } = new ErrorService();

    /// <summary>The process-wide on-disk archive of recent runs (inputs,
    /// output, error, manifest). Wired by the add-in's <c>AutoOpen</c>;
    /// <c>PyRun.Execute*</c> writes here on every run when a
    /// <c>RunArchiveContext</c> is supplied.</summary>
    public static RunArchive RunArchive { get; set; } = BuildDefaultRunArchive();

    /// <summary>Strategy for "what workbook is active right now".</summary>
    public static IWorkbookContext WorkbookContext { get; set; } = NullWorkbookContext.Instance;

    /// <summary>Per-workbook result of the project-structure check, and the single
    /// readiness gate every ribbon control reads. The COM event sink fills it on
    /// workbook open and activate for enabled workbooks (and Enable/Repair and the
    /// Run guard refresh it); the ribbon turns it — via
    /// <see cref="HealthRegistry.ReadinessOf"/> — into the enabled-state of both the
    /// data controls and the Enable/Repair button, a cheap in-memory lookup with no
    /// per-render I/O.</summary>
    public static HealthRegistry Health { get; set; } = new HealthRegistry();

    /// <summary>
    /// Hook the ribbon registers (in <c>RibbonOnLoad</c>) so non-ribbon
    /// components can ask the ribbon to repaint. The motivating caller is
    /// the COM event sink on <c>WorkbookActivate</c>: the active workbook
    /// key changed, so every getter now renders a different state, but no
    /// <see cref="StateService.StateChanged"/> fired because nothing in the
    /// registry mutated.
    ///
    /// <para>The ribbon's implementation queues <c>IRibbonUI.Invalidate</c>
    /// onto Excel's macro thread, so callers may invoke this from any
    /// thread. It is <see langword="null"/> until the ribbon registers it
    /// (and after the add-in unloads), so callers invoke it null-conditionally
    /// — a no-op in that window.</para>
    /// </summary>
    public static Action? RequestRibbonInvalidate { get; set; }

    /// <summary>
    /// Hook the COM event sink registers so the ribbon can ask it to (re)point
    /// the live <see cref="ScriptDirectoryWatcher"/> at the active workbook's
    /// <c>userScripts</c> folder. The motivating caller is <c>OnEnablePyExcel</c>:
    /// Enable provisions that folder but fires no <c>WorkbookActivate</c>, so the
    /// watcher would otherwise not start until the next activation.
    ///
    /// <para><see langword="null"/> until the sink registers it (and after the
    /// add-in unloads), so callers invoke it null-conditionally.</para>
    /// </summary>
    public static Action? RequestScriptRefresh { get; set; }

    /// <summary>
    /// Build the default <see cref="RunArchive"/> rooted under the per-user
    /// local app-data folder (<c>%LOCALAPPDATA%\PyExcel\runs</c> on Windows;
    /// the XDG-derived equivalent on Linux). Falls back to the system temp
    /// dir when the platform doesn't have a local app-data folder — keeps
    /// the service constructable on every CI lane.
    /// </summary>
    private static RunArchive BuildDefaultRunArchive()
    {
        var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var root = string.IsNullOrEmpty(localAppData)
            ? Path.Combine(Path.GetTempPath(), "PyExcel", "runs")
            : Path.Combine(localAppData, "PyExcel", "runs");
        return new RunArchive(root);
    }
}

/// <summary>
/// Process-wide registry of each enabled workbook's last project-structure check
/// (see <see cref="ProjectStructureValidator"/>). Thread-safe: the COM sink writes
/// from workbook events while the ribbon reads from getEnabled callbacks.
/// </summary>
public sealed class HealthRegistry
{
    private readonly ConcurrentDictionary<string, ProjectStructureCheck> _checks =
        new(StringComparer.Ordinal);

    /// <summary>Record the latest structure check for a workbook.</summary>
    public void Set(string workbookKey, ProjectStructureCheck check)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        if (check is null) throw new ArgumentNullException(nameof(check));
        _checks[workbookKey] = check;
    }

    /// <summary>The last check for a workbook, or null if none was recorded.</summary>
    public ProjectStructureCheck? Get(string? workbookKey)
        => workbookKey is not null && _checks.TryGetValue(workbookKey, out var c) ? c : null;

    /// <summary>The consolidated readiness verdict for a workbook, from its
    /// <paramref name="enabled"/> flag and the last recorded structure check — the
    /// single gate the ribbon reads. Data controls are live iff
    /// <see cref="ProjectReadiness.Ready"/>; the Enable button (which doubles as
    /// Repair) is live for every other value. See
    /// <see cref="ProjectReadinessClassifier.Classify"/> for the rule.</summary>
    public ProjectReadiness ReadinessOf(string? workbookKey, bool enabled)
        => ProjectReadinessClassifier.Classify(enabled, Get(workbookKey));

    /// <summary>Forget a workbook's check (on close, or when it's no longer a
    /// project).</summary>
    public void Clear(string workbookKey)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        _checks.TryRemove(workbookKey, out _);
    }
}
