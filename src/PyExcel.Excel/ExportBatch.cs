using System;
using System.Collections.Generic;

namespace PyExcel.Excel;

/// <summary>One row of the Export Wizard: a source range to write to a
/// target file. The fully-resolved <see cref="ExportPlan"/> is produced
/// per job at run time via <see cref="ExportPlanner"/>.</summary>
public sealed record ExportJob(string SourceRange, string TargetPath);

/// <summary>The outcome of validating the Export Wizard's rows.</summary>
public sealed class ExportBatchValidationResult
{
    private ExportBatchValidationResult(bool isValid, string? error, IReadOnlyList<ExportJob> jobs)
    {
        IsValid = isValid;
        ErrorMessage = error;
        Jobs = jobs;
    }

    public bool IsValid { get; }
    public string? ErrorMessage { get; }

    /// <summary>The validated jobs (trimmed) when <see cref="IsValid"/>;
    /// empty otherwise.</summary>
    public IReadOnlyList<ExportJob> Jobs { get; }

    internal static ExportBatchValidationResult Ok(IReadOnlyList<ExportJob> jobs)
        => new(true, null, jobs);

    internal static ExportBatchValidationResult Fail(string error)
        => new(false, error, Array.Empty<ExportJob>());
}

/// <summary>
/// Pure validation for the Export Wizard (the Phase 8 port of v1's
/// <c>frmExportWizard</c>, reshaped to v2's CSV/TSV export). Each row is
/// validated through the same <see cref="ExportPlanner"/> the single-export
/// path uses, so a row is rejected for exactly the reasons a single export
/// would be (blank range, blank/unsupported target). Kept cross-platform
/// so it's unit-tested on Linux CI.
/// </summary>
public static class ExportBatchValidator
{
    /// <summary>Validate every row, returning the trimmed jobs to run or
    /// the first row's error (1-based row number in the message).</summary>
    public static ExportBatchValidationResult Validate(
        IReadOnlyList<ExportJob> jobs, string? workbookDirectory)
    {
        if (jobs is null) throw new ArgumentNullException(nameof(jobs));
        if (jobs.Count == 0)
            return ExportBatchValidationResult.Fail("Add at least one export row.");

        var validated = new List<ExportJob>(jobs.Count);
        var seenTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < jobs.Count; i++)
        {
            var job = jobs[i];
            try
            {
                // Reuse the planner as the per-row validity check; it
                // throws FormatException for blank/unsupported fields.
                ExportPlanner.Create(job.SourceRange, job.TargetPath, workbookDirectory);
            }
            catch (FormatException ex)
            {
                return ExportBatchValidationResult.Fail($"Row {i + 1}: {ex.Message}");
            }
            catch (ArgumentException ex)
            {
                return ExportBatchValidationResult.Fail($"Row {i + 1}: {ex.Message}");
            }

            var source = job.SourceRange.Trim();
            var target = job.TargetPath.Trim();
            if (!seenTargets.Add(target))
                return ExportBatchValidationResult.Fail(
                    $"Row {i + 1}: two rows write to the same file '{target}'.");

            validated.Add(new ExportJob(source, target));
        }

        return ExportBatchValidationResult.Ok(validated);
    }
}
