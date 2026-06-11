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

/// <summary>Whether a list's orientation is already determined by the
/// target range, or must be asked of the user.</summary>
public readonly struct OrientationResolution
{
    private OrientationResolution(bool ask, ListOrientation orientation)
    {
        Ask = ask;
        Orientation = orientation;
    }

    /// <summary>True when the target is a single cell, so orientation is
    /// ambiguous and the user must choose.</summary>
    public bool Ask { get; }

    /// <summary>The auto-detected orientation when <see cref="Ask"/> is
    /// false; meaningless otherwise.</summary>
    public ListOrientation Orientation { get; }

    internal static OrientationResolution Resolved(ListOrientation o) => new(false, o);
    internal static OrientationResolution NeedsPrompt() => new(true, default);
}

/// <summary>
/// Decides how a 1-D list spills into a target range — the v1 paste-
/// direction rule, lifted out of the COM path so it's unit-tested on
/// Linux CI. A multi-cell target dictates the orientation by its shape
/// (wider-or-square ⇒ horizontal, taller ⇒ vertical); a single cell is
/// ambiguous and the caller must prompt (the <c>OrientationForm</c>).
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

        // Multi-cell — match v1: paste across when the block is at least as
        // wide as it is tall, otherwise down.
        return OrientationResolution.Resolved(
            targetCols >= targetRows ? ListOrientation.Horizontal : ListOrientation.Vertical);
    }
}
