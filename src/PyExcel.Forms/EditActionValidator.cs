using System;
using System.Collections.Generic;
using PyExcel.State;

namespace PyExcel.Forms;

/// <summary>
/// The outcome of validating the EditAction dialog's fields: either a
/// built <see cref="RibbonAction"/> ready to hand to
/// <c>StateService.AddAction</c>, or an inline error message for the
/// dialog to show without closing.
/// </summary>
public sealed class EditActionValidationResult
{
    private EditActionValidationResult(bool isValid, string? error, RibbonAction? action)
    {
        IsValid = isValid;
        ErrorMessage = error;
        Action = action;
    }

    /// <summary>True when the fields produced a valid action.</summary>
    public bool IsValid { get; }

    /// <summary>The inline error to show, or null when <see cref="IsValid"/>.</summary>
    public string? ErrorMessage { get; }

    /// <summary>The built action, or null when not <see cref="IsValid"/>.</summary>
    public RibbonAction? Action { get; }

    internal static EditActionValidationResult Ok(RibbonAction action)
        => new(true, null, action);

    internal static EditActionValidationResult Fail(string error)
        => new(false, error, null);
}

/// <summary>
/// Pure validation for the Add/Edit-action dialog — the logic the
/// WinForms <c>EditActionForm</c> runs on Save before it closes, kept
/// here so it's unit-testable on Linux CI without a desktop.
///
/// <para>The dialog catches an invalid result and shows
/// <see cref="EditActionValidationResult.ErrorMessage"/> inline; only a
/// valid result closes the form. This satisfies the Phase 8 exit
/// criterion that invalid input is caught in the dialog, not downstream.</para>
/// </summary>
public static class EditActionValidator
{
    /// <summary>
    /// Validate the dialog fields and, on success, build the
    /// <see cref="RibbonAction"/>.
    /// </summary>
    /// <param name="name">The action name the user typed.</param>
    /// <param name="script">The script the user selected.</param>
    /// <param name="input">The input range reference.</param>
    /// <param name="output">The output range reference.</param>
    /// <param name="kwargs">Optional parsed keyword arguments (already
    /// validated by <see cref="KwargsText.TryParse"/>); pass null/empty
    /// for none.</param>
    /// <param name="existingActionNames">The names already in the
    /// workbook, used to reject a duplicate.</param>
    /// <param name="originalName">When editing an existing action, its
    /// current name — so renaming the action onto its own name (i.e.
    /// leaving the name alone) is not flagged as a duplicate. Pass null
    /// in Add mode.</param>
    /// <param name="keepOutputOpen">Whether a successful run leaves the
    /// run-output window open (default true). Carried straight onto the
    /// built <see cref="RibbonAction"/>; not otherwise validated.</param>
    public static EditActionValidationResult Validate(
        string? name,
        string? script,
        string? input,
        string? output,
        IReadOnlyDictionary<string, string>? kwargs,
        IEnumerable<string> existingActionNames,
        string? originalName = null,
        bool keepOutputOpen = true)
    {
        if (existingActionNames is null)
            throw new ArgumentNullException(nameof(existingActionNames));

        var trimmedName = (name ?? string.Empty).Trim();
        if (trimmedName.Length == 0)
            return EditActionValidationResult.Fail("Enter a name for the action.");

        // Names are compared the same way StateService.AddAction upserts
        // them — ordinal, case-sensitive. An edit that keeps the name is
        // allowed; only a collision with a *different* action is rejected.
        foreach (var existing in existingActionNames)
        {
            if (string.Equals(existing, trimmedName, StringComparison.Ordinal) &&
                !string.Equals(existing, originalName, StringComparison.Ordinal))
            {
                return EditActionValidationResult.Fail(
                    $"An action named '{trimmedName}' already exists. " +
                    "Choose a different name.");
            }
        }

        var trimmedScript = (script ?? string.Empty).Trim();
        if (trimmedScript.Length == 0)
            return EditActionValidationResult.Fail("Select a script for this action.");

        var trimmedInput = (input ?? string.Empty).Trim();
        if (trimmedInput.Length == 0)
            return EditActionValidationResult.Fail(
                "Enter the input range (e.g. A1:B10, or Sheet1!A1:B10).");

        var trimmedOutput = (output ?? string.Empty).Trim();
        if (trimmedOutput.Length == 0)
            return EditActionValidationResult.Fail(
                "Enter the output range (e.g. D1, or Sheet1!D1).");

        // Copy the kwargs into a fresh dictionary so the result doesn't
        // alias the caller's instance, and drop the empty case to null to
        // match RibbonAction's "no kwargs" convention.
        IReadOnlyDictionary<string, string>? actionKwargs = null;
        if (kwargs is not null && kwargs.Count > 0)
        {
            var copy = new Dictionary<string, string>(kwargs.Count, StringComparer.Ordinal);
            foreach (var kv in kwargs) copy[kv.Key] = kv.Value;
            actionKwargs = copy;
        }

        var action = new RibbonAction(
            trimmedName, trimmedScript, trimmedInput, trimmedOutput, actionKwargs,
            keepOutputOpen);
        return EditActionValidationResult.Ok(action);
    }
}
