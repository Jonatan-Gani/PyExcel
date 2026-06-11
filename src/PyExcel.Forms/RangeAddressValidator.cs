using System;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// The outcome of validating a range-picker entry: either a single,
/// well-formed range address, or an inline error.
/// </summary>
public sealed class RangeAddressValidationResult
{
    private RangeAddressValidationResult(bool isValid, string? error, string? address)
    {
        IsValid = isValid;
        ErrorMessage = error;
        Address = address;
    }

    public bool IsValid { get; }
    public string? ErrorMessage { get; }

    /// <summary>The trimmed range address — null when not
    /// <see cref="IsValid"/>.</summary>
    public string? Address { get; }

    internal static RangeAddressValidationResult Ok(string address)
        => new(true, null, address);

    internal static RangeAddressValidationResult Fail(string error)
        => new(false, error, null);
}

/// <summary>
/// Pure validation for the range-picker dialog. A picked range is a
/// <em>single</em> range reference (optionally sheet-qualified), so it's
/// validated through <see cref="RibbonRangeParser"/> — the same parser the
/// run dispatcher uses — and rejected if it carries a name binding or
/// resolves to more than one range. Kept cross-platform so the rule is
/// unit-tested on Linux CI.
/// </summary>
public static class RangeAddressValidator
{
    /// <summary>Validate that <paramref name="address"/> is exactly one
    /// plain range reference, returning it trimmed.</summary>
    public static RangeAddressValidationResult Validate(string? address)
    {
        var trimmed = (address ?? string.Empty).Trim();
        if (trimmed.Length == 0)
            return RangeAddressValidationResult.Fail("Enter a range (e.g. A1, or Sheet1!A1:C10).");

        // The picker yields one range, not the ribbon's multi-binding
        // syntax — reject ';' (several ranges) and '=' (a name binding).
        if (trimmed.IndexOf(';') >= 0)
            return RangeAddressValidationResult.Fail(
                "Enter a single range — remove the ';'.");
        if (trimmed.IndexOf('=') >= 0)
            return RangeAddressValidationResult.Fail(
                "Enter a plain range without a 'name=' prefix.");

        try
        {
            var bindings = RibbonRangeParser.Parse(trimmed);
            if (bindings.Count != 1)
                return RangeAddressValidationResult.Fail("Enter exactly one range.");
            return RangeAddressValidationResult.Ok(bindings[0].RangeText);
        }
        catch (FormatException ex)
        {
            return RangeAddressValidationResult.Fail(ex.Message);
        }
    }
}
