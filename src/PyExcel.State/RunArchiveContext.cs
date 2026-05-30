using System;

namespace PyExcel.State;

/// <summary>
/// What <see cref="PyRun"/> needs to archive a run on top of what it
/// already knows (script, inputs, output, error): a destination
/// <see cref="RunArchive"/>, a <see cref="Source"/> label describing
/// which surface invoked the run, and the <see cref="WorkbookKey"/> of
/// the active workbook (if any).
///
/// <para>Passing this is opt-in — the existing <c>PyRun.Execute*</c>
/// overloads default it to <see langword="null"/> so unit tests and
/// callers that don't want archiving aren't forced to set one up.</para>
/// </summary>
public sealed class RunArchiveContext
{
    /// <summary>Where to write the archive.</summary>
    public RunArchive Archive { get; }

    /// <summary>Origin label captured on every archived run. Mirrors the
    /// <c>Source</c> field on <see cref="KernelErrorRecord"/>.</summary>
    public string Source { get; }

    /// <summary>Active workbook key at the time of the run, or
    /// <see langword="null"/> if none is bound (kernel boot, an
    /// unattached UDF call).</summary>
    public string? WorkbookKey { get; }

    public RunArchiveContext(RunArchive archive, string source, string? workbookKey)
    {
        Archive = archive ?? throw new ArgumentNullException(nameof(archive));
        if (string.IsNullOrWhiteSpace(source))
            throw new ArgumentException("source must be a non-empty label", nameof(source));
        Source = source;
        WorkbookKey = workbookKey;
    }
}
