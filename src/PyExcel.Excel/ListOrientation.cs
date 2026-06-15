using System;

namespace PyExcel.Excel;

/// <summary>Which way a 1-D list spills into the sheet.</summary>
public enum ListOrientation
{
    /// <summary>Across a row (one row, N columns).</summary>
    Horizontal,

    /// <summary>Down a column (N rows, one column).</summary>
    Vertical,
}

/// <summary>What the caller must do to spill a 1-D list into a target range.</summary>
public enum OrientationDecision
{
    /// <summary>The target is a single cell — the direction is ambiguous, so the
    /// caller must ask the user (the <c>OrientationForm</c>).</summary>
    Ask,

    /// <summary>The target is a single row or single column, which fixes the
    /// direction (<see cref="OrientationResolution.Orientation"/>).</summary>
    Resolved,

    /// <summary>The target is a 2-D block (more than one row AND more than one
    /// column) — a 1-D list can't fill it unambiguously, so the caller must reject
    /// the write.</summary>
    Invalid,
}

/// <summary>How a 1-D list's orientation resolves against its target range:
/// already determined by the range's shape, ambiguous (must ask the user), or
/// invalid (a 2-D block).</summary>
public readonly struct OrientationResolution
{
    private OrientationResolution(OrientationDecision decision, ListOrientation orientation)
    {
        Decision = decision;
        Orientation = orientation;
    }

    /// <summary>What the caller must do with this result.</summary>
    public OrientationDecision Decision { get; }

    /// <summary>The auto-detected orientation when <see cref="Decision"/> is
    /// <see cref="OrientationDecision.Resolved"/>; meaningless otherwise.</summary>
    public ListOrientation Orientation { get; }

    /// <summary>True when the target is a single cell, so the caller must prompt.</summary>
    public bool Ask => Decision == OrientationDecision.Ask;

    /// <summary>True when the target is a 2-D block a 1-D list can't fill.</summary>
    public bool IsInvalid => Decision == OrientationDecision.Invalid;

    internal static OrientationResolution Resolved(ListOrientation o)
        => new(OrientationDecision.Resolved, o);
    internal static OrientationResolution NeedsPrompt()
        => new(OrientationDecision.Ask, default);
    internal static OrientationResolution Invalid()
        => new(OrientationDecision.Invalid, default);
}

/// <summary>
/// Decides how a 1-D list spills into a target range, lifted out of the COM path
/// so it's unit-tested on Linux CI. The rule, by the target's shape:
/// <list type="bullet">
///   <item>single cell (≤ 1 × ≤ 1) ⇒ <see cref="OrientationDecision.Ask"/> — the
///     caller prompts via the <c>OrientationForm</c>;</item>
///   <item>a single row (1 × N) ⇒ horizontal; a single column (N × 1) ⇒ vertical
///     (both <see cref="OrientationDecision.Resolved"/>);</item>
///   <item>a 2-D block (rows &gt; 1 AND cols &gt; 1) ⇒
///     <see cref="OrientationDecision.Invalid"/> — a list can't fill a block
///     unambiguously, so the caller rejects the write.</item>
/// </list>
/// </summary>
public static class OrientationResolver
{
    /// <summary>Resolve from the target range's dimensions.</summary>
    /// <param name="targetRows">Row count of the target range (≥ 0).</param>
    /// <param name="targetCols">Column count of the target range (≥ 0).</param>
    public static OrientationResolution Resolve(int targetRows, int targetCols)
    {
        // Single cell (or degenerate) — ambiguous, ask the user.
        if (targetRows <= 1 && targetCols <= 1)
            return OrientationResolution.NeedsPrompt();

        // A single row or single column dictates the direction.
        if (targetRows <= 1) return OrientationResolution.Resolved(ListOrientation.Horizontal);
        if (targetCols <= 1) return OrientationResolution.Resolved(ListOrientation.Vertical);

        // Both dimensions > 1 — a 2-D block can't take a 1-D list.
        return OrientationResolution.Invalid();
    }
}
