using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// The Python type a range binding declares. This is the user-facing
/// vocabulary of the typed I/O contract — see
/// <c>docs/typed-io-contract.md</c> for the full coercion matrix.
///
/// <para>Deliberately distinct from <c>PyExcel.Excel.ArrowShape</c>. A
/// shape describes the physical layout of an Arrow buffer (the host needs
/// it to rebuild the value); a declared type is the contract the kernel
/// constructs against. The two legitimately disagree mid-coercion — a
/// range declared <see cref="Set"/> travels as a plain grid — and
/// <c>ArrowShape</c> carries producer-only outcomes (<c>Chart</c>,
/// <c>Image</c>) that must never be selectable as an input declaration.</para>
///
/// <para>The wire names in <see cref="PyExcelTypes.WireName"/> are mirrored
/// by <c>embedded/pyexcel/kernel/types.py</c>; changing one requires
/// changing the other.</para>
/// </summary>
public enum PyExcelType
{
    /// <summary>No declaration — resolved from the range's dimensions at
    /// run time by <see cref="PyExcelTypes.ResolveAuto"/>. Never travels
    /// to the kernel on an input; on an output it means "do not enforce".</summary>
    Auto = 0,

    /// <summary>A <c>pandas.DataFrame</c>; the first row supplies column names.</summary>
    DataFrame,

    /// <summary>A <c>pandas.Series</c>; the first cell supplies the name.</summary>
    Series,

    /// <summary>A Python <c>list</c>; every cell is data.</summary>
    List,

    /// <summary>A Python <c>tuple</c>; every cell is data.</summary>
    Tuple,

    /// <summary>A Python <c>set</c> of the distinct cell values.</summary>
    Set,

    /// <summary>A Python <c>dict</c>; two columns read as key to value,
    /// three or more as column-oriented lists keyed by the header row.</summary>
    Dict,

    /// <summary>A <c>numpy.ndarray</c> matching the range's dimensions.</summary>
    NDArray,

    /// <summary>A single Python scalar; only a one-cell range is valid.</summary>
    Scalar,
}

/// <summary>
/// Wire names, parsing, and the <see cref="PyExcelType.Auto"/> resolution
/// rule for <see cref="PyExcelType"/>. Pure logic — no COM, no Arrow.
/// </summary>
public static class PyExcelTypes
{
    /// <summary>Every type in dropdown order, <see cref="PyExcelType.Auto"/>
    /// first because it is the default for a new binding.</summary>
    public static readonly IReadOnlyList<PyExcelType> All = new[]
    {
        PyExcelType.Auto,
        PyExcelType.DataFrame,
        PyExcelType.Series,
        PyExcelType.List,
        PyExcelType.Tuple,
        PyExcelType.Set,
        PyExcelType.Dict,
        PyExcelType.NDArray,
        PyExcelType.Scalar,
    };

    /// <summary>
    /// The lowercase token used in the binding grammar and on the kernel
    /// wire. Mirrored by <c>embedded/pyexcel/kernel/types.py</c>.
    /// </summary>
    public static string WireName(PyExcelType type) => type switch
    {
        PyExcelType.Auto => "auto",
        PyExcelType.DataFrame => "dataframe",
        PyExcelType.Series => "series",
        PyExcelType.List => "list",
        PyExcelType.Tuple => "tuple",
        PyExcelType.Set => "set",
        PyExcelType.Dict => "dict",
        PyExcelType.NDArray => "ndarray",
        PyExcelType.Scalar => "scalar",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "unknown declared type"),
    };

    /// <summary>
    /// The label shown in the form's type box. Differs from
    /// <see cref="WireName"/> only in casing.
    /// </summary>
    public static string DisplayName(PyExcelType type) => type switch
    {
        PyExcelType.Auto => "Auto",
        PyExcelType.DataFrame => "DataFrame",
        PyExcelType.Series => "Series",
        PyExcelType.List => "List",
        PyExcelType.Tuple => "Tuple",
        PyExcelType.Set => "Set",
        PyExcelType.Dict => "Dict",
        PyExcelType.NDArray => "NDArray",
        PyExcelType.Scalar => "Scalar",
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, "unknown declared type"),
    };

    /// <summary>
    /// Parse a wire/display token case-insensitively. Returns
    /// <see langword="false"/> for anything unrecognised — the binding
    /// grammar relies on that to tell a declared type from a name that
    /// merely happens to contain a colon.
    /// </summary>
    public static bool TryParse(string? token, out PyExcelType type)
    {
        type = PyExcelType.Auto;
        if (string.IsNullOrWhiteSpace(token)) return false;

        switch (token!.Trim().ToLowerInvariant())
        {
            case "auto": type = PyExcelType.Auto; return true;
            case "dataframe": type = PyExcelType.DataFrame; return true;
            case "series": type = PyExcelType.Series; return true;
            case "list": type = PyExcelType.List; return true;
            case "tuple": type = PyExcelType.Tuple; return true;
            case "set": type = PyExcelType.Set; return true;
            case "dict": type = PyExcelType.Dict; return true;
            case "ndarray": type = PyExcelType.NDArray; return true;
            case "scalar": type = PyExcelType.Scalar; return true;
            default: return false;
        }
    }

    /// <summary>
    /// Resolve <see cref="PyExcelType.Auto"/> against a range's measured
    /// dimensions: a single cell is a scalar, a single row or column is a
    /// list, anything larger is a DataFrame. Any already-declared type is
    /// returned untouched.
    ///
    /// <para>This runs at run time, in the run driver, because that is the
    /// only place the range has actually been measured — the action dialog
    /// holds nothing but an address string.</para>
    /// </summary>
    /// <param name="declared">The binding's declared type.</param>
    /// <param name="rows">Row count of the resolved range (at least 1).</param>
    /// <param name="columns">Column count of the resolved range (at least 1).</param>
    public static PyExcelType ResolveAuto(PyExcelType declared, int rows, int columns)
    {
        if (declared != PyExcelType.Auto) return declared;
        if (rows <= 1 && columns <= 1) return PyExcelType.Scalar;
        if (rows <= 1 || columns <= 1) return PyExcelType.List;
        return PyExcelType.DataFrame;
    }
}
