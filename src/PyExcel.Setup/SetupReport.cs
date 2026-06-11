using System;
using System.Linq;
using System.Text;

namespace PyExcel.Setup;

/// <summary>
/// Formats a <see cref="SetupResult"/> into the human-readable summary the
/// first-run wizard appends below the live venv/pip log. Pure string work,
/// split out of the WinForms form so it's unit-testable on Linux.
/// </summary>
public static class SetupReport
{
    /// <summary>A one-line headline for the run: success, or the first stage
    /// that failed and why.</summary>
    public static string Headline(SetupResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));
        if (result.Success) return "Setup completed successfully.";

        var failed = result.Steps.FirstOrDefault(s => !s.Success);
        if (failed is null) return "Setup failed.";
        return string.IsNullOrWhiteSpace(failed.FailureReason)
            ? $"Setup failed at '{failed.Name}'."
            : $"Setup failed at '{failed.Name}': {failed.FailureReason}";
    }

    /// <summary>A per-stage transcript — one <c>[ok] / [fail]</c> line per
    /// step, with the failure reason inline — followed by the
    /// <see cref="Headline"/>.</summary>
    public static string Summarize(SetupResult result)
    {
        if (result is null) throw new ArgumentNullException(nameof(result));

        var sb = new StringBuilder();
        foreach (var step in result.Steps)
        {
            if (step.Success)
            {
                sb.Append("[ok]   ").AppendLine(step.Name);
            }
            else
            {
                sb.Append("[fail] ").Append(step.Name);
                if (!string.IsNullOrWhiteSpace(step.FailureReason))
                    sb.Append(" — ").Append(step.FailureReason);
                sb.AppendLine();
            }
        }
        sb.AppendLine();
        sb.Append(Headline(result));
        return sb.ToString();
    }
}
