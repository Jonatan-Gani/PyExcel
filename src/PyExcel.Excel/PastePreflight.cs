using System;

namespace PyExcel.Excel;

/// <summary>
/// Pure-logic preflight for the Paste operation. The COM-bound
/// <c>PasteService</c> calls these helpers to determine
///   (a) the cell-footprint a decoded payload will occupy, and
///   (b) whether the target range already holds content the paste would
///       overwrite (the signal for the destructive-action confirmation
///       dialog).
///
/// <para>Cross-platform on purpose: no Excel-DNA dependency, no COM. The
/// caller is responsible for stripping Excel-DNA's <c>ExcelEmpty</c> /
/// <c>ExcelMissing</c> sentinels from the <c>Value2</c> snapshot before
/// invoking <see cref="RangeHasContent"/> — those types are net48-only
/// and don't belong in a shared planner.</para>
/// </summary>
public static class PastePreflight
{
    /// <summary>How many rows × columns the decoded payload will occupy
    /// when written into the target range. Mirrors <c>PasteService</c>'s
    /// write semantics: a 2-D array uses its own dimensions, a 1-D array
    /// spills as one row, a scalar lands in a single cell, null returns
    /// the zero footprint so the service short-circuits.</summary>
    public static (int rows, int cols) Footprint(object? decoded)
    {
        if (decoded is null) return (0, 0);
        return decoded switch
        {
            object?[,] table => (table.GetLength(0), table.GetLength(1)),
            object?[] vector => (1, vector.Length),
            _ => (1, 1),
        };
    }

    /// <summary>Does the target snapshot hold anything the paste would
    /// overwrite? <c>null</c> and empty strings count as "no content";
    /// everything else (numbers, dates, non-empty strings, errors)
    /// counts as occupied. Excel-DNA sentinels are <em>not</em>
    /// recognised — the caller must strip them first.</summary>
    public static bool RangeHasContent(object? value2Snapshot)
    {
        if (value2Snapshot is null) return false;

        // Pattern annotations match the codebase convention (see
        // PasteService.WriteToRange): `object?[,]` and `object?[]` are
        // the same runtime types as the non-nullable variants — the
        // annotation is pure compile-time metadata — but it lets the
        // indexer return `object?` so cells that are legitimately null
        // (COM routinely hands those back) flow into CellHasContent
        // without a null-suppression warning.
        if (value2Snapshot is object?[,] arr)
        {
            int r0 = arr.GetLowerBound(0);
            int c0 = arr.GetLowerBound(1);
            int height = arr.GetLength(0);
            int width = arr.GetLength(1);
            for (int i = 0; i < height; i++)
                for (int j = 0; j < width; j++)
                    if (CellHasContent(arr[r0 + i, c0 + j])) return true;
            return false;
        }

        if (value2Snapshot is object?[] vec)
        {
            for (int i = 0; i < vec.Length; i++)
                if (CellHasContent(vec[i])) return true;
            return false;
        }

        return CellHasContent(value2Snapshot);
    }

    private static bool CellHasContent(object? cell)
    {
        if (cell is null) return false;
        if (cell is string s) return s.Length > 0;
        return true;
    }
}
