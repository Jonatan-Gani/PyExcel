namespace PyExcel.State;

/// <summary>
/// The single readiness verdict for a workbook, derived from whether it has been
/// enabled for PyExcel and whether its on-disk project structure
/// (<see cref="ProjectStructureValidator"/>) currently validates.
///
/// <para>Every ribbon control reads its enabled-state from this one classification,
/// so the gate the user sees is computed in exactly one place: the data controls
/// (Run, Edit, the Script / Input / Output / Actions fields, Import, Export, Paste)
/// are live iff <see cref="Ready"/>, and the Enable button — which doubles as Repair
/// — is live iff <em>not</em> <see cref="Ready"/>. The two are exact complements, so
/// they can never drift out of step.</para>
/// </summary>
public enum ProjectReadiness
{
    /// <summary>The active workbook hasn't been enabled for PyExcel — only the
    /// Enable button is actionable.</summary>
    NotEnabled,

    /// <summary>Enabled, but the project structure is missing / incomplete (or has
    /// not validated yet): the data controls stay off and Enable doubles as Repair
    /// until the environment is confirmed whole.</summary>
    NeedsRepair,

    /// <summary>Enabled and the structure validated healthy — every data control is
    /// live and the elements (scripts, actions) are loaded.</summary>
    Ready,
}

/// <summary>
/// Pure classification of a workbook's readiness from its enabled flag and last
/// structure check. Kept separate from <see cref="HealthRegistry"/> and the ribbon
/// so the gate every control depends on is unit-testable without any COM / Excel
/// plumbing.
/// </summary>
public static class ProjectReadinessClassifier
{
    /// <summary>
    /// Classify readiness. A workbook is <see cref="ProjectReadiness.Ready"/> only
    /// when it is enabled <b>and</b> its latest structure check passed; an enabled
    /// workbook with a failed — or not-yet-recorded — check is
    /// <see cref="ProjectReadiness.NeedsRepair"/> so the data controls stay off (and
    /// the elements stay unloaded) until the environment is positively confirmed.
    /// Treating "no check yet" as not-ready is the conservative, self-healing choice:
    /// the next validate (open / activate / Run) or a Repair click records a verdict
    /// and flips the gate the moment the environment is whole.
    /// </summary>
    public static ProjectReadiness Classify(bool enabled, ProjectStructureCheck? check)
    {
        if (!enabled) return ProjectReadiness.NotEnabled;
        return check is not null && check.Ok
            ? ProjectReadiness.Ready
            : ProjectReadiness.NeedsRepair;
    }
}
