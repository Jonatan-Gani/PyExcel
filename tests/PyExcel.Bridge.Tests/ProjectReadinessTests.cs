using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// The consolidated readiness gate: one verdict, derived from a workbook's enabled
/// flag and its last structure check, that drives both the data controls (live iff
/// <see cref="ProjectReadiness.Ready"/>) and the Enable/Repair button (live
/// otherwise). These tests pin the classification so the ribbon and the COM sink,
/// which both read it, can't drift.
/// </summary>
public class ProjectReadinessTests
{
    private static ProjectStructureCheck Broken =>
        new(false, new[] { "the Python virtual environment (.pyexcel-venv)" });

    // -- Classify: the pure rule ------------------------------------------------

    [Fact]
    public void Classify_NotEnabled_RegardlessOfCheck()
    {
        Assert.Equal(ProjectReadiness.NotEnabled,
            ProjectReadinessClassifier.Classify(enabled: false, check: null));
        Assert.Equal(ProjectReadiness.NotEnabled,
            ProjectReadinessClassifier.Classify(enabled: false, ProjectStructureCheck.Healthy));
        Assert.Equal(ProjectReadiness.NotEnabled,
            ProjectReadinessClassifier.Classify(enabled: false, Broken));
    }

    [Fact]
    public void Classify_Ready_WhenEnabledAndHealthy()
    {
        Assert.Equal(ProjectReadiness.Ready,
            ProjectReadinessClassifier.Classify(enabled: true, ProjectStructureCheck.Healthy));
    }

    [Fact]
    public void Classify_NeedsRepair_WhenEnabledAndStructureBroken()
    {
        Assert.Equal(ProjectReadiness.NeedsRepair,
            ProjectReadinessClassifier.Classify(enabled: true, Broken));
    }

    [Fact]
    public void Classify_NeedsRepair_WhenEnabledButNotYetChecked()
    {
        // The conservative, self-healing default: an enabled workbook with no recorded
        // check is treated as not-ready, so the data controls stay off until a validate
        // (or Repair) records a positive verdict — never optimistically enabling Run on
        // an unverified environment.
        Assert.Equal(ProjectReadiness.NeedsRepair,
            ProjectReadinessClassifier.Classify(enabled: true, check: null));
    }

    // -- ReadinessOf: reading the registry --------------------------------------

    [Fact]
    public void ReadinessOf_ReflectsRecordedHealthyCheck()
    {
        var health = new HealthRegistry();
        health.Set("wb", ProjectStructureCheck.Healthy);

        Assert.Equal(ProjectReadiness.Ready, health.ReadinessOf("wb", enabled: true));
        // Even with a healthy check on file, a not-enabled workbook is NotEnabled.
        Assert.Equal(ProjectReadiness.NotEnabled, health.ReadinessOf("wb", enabled: false));
    }

    [Fact]
    public void ReadinessOf_ReflectsRecordedBrokenCheck()
    {
        var health = new HealthRegistry();
        health.Set("wb", Broken);

        Assert.Equal(ProjectReadiness.NeedsRepair, health.ReadinessOf("wb", enabled: true));
    }

    [Fact]
    public void ReadinessOf_NoRecord_EnabledIsNeedsRepair_DisabledIsNotEnabled()
    {
        var health = new HealthRegistry();

        Assert.Equal(ProjectReadiness.NeedsRepair, health.ReadinessOf("never-checked", enabled: true));
        Assert.Equal(ProjectReadiness.NotEnabled, health.ReadinessOf("never-checked", enabled: false));
    }

    [Fact]
    public void ReadinessOf_NullKey_DerivesFromEnabledFlagOnly()
    {
        var health = new HealthRegistry();

        Assert.Equal(ProjectReadiness.NotEnabled, health.ReadinessOf(null, enabled: false));
        Assert.Equal(ProjectReadiness.NeedsRepair, health.ReadinessOf(null, enabled: true));
    }

    [Fact]
    public void ReadinessOf_AfterClear_FallsBackToUnchecked()
    {
        var health = new HealthRegistry();
        health.Set("wb", ProjectStructureCheck.Healthy);
        Assert.Equal(ProjectReadiness.Ready, health.ReadinessOf("wb", enabled: true));

        // Clearing (workbook closed, or no longer a project) drops the verdict, so an
        // enabled workbook reverts to NeedsRepair until re-validated.
        health.Clear("wb");
        Assert.Equal(ProjectReadiness.NeedsRepair, health.ReadinessOf("wb", enabled: true));
    }
}
