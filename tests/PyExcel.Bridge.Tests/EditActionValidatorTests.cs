using System;
using System.Collections.Generic;
using PyExcel.Forms;
using PyExcel.State;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class EditActionValidatorTests
{
    private static readonly string[] NoExisting = Array.Empty<string>();

    private static IReadOnlyDictionary<string, string>? NoKwargs => null;

    // -------------------------------------------------------------------------
    // Happy path
    // -------------------------------------------------------------------------

    [Fact]
    public void Validate_AllFields_BuildsAction()
    {
        var result = EditActionValidator.Validate(
            "Forecast", "model.py", "A1:B10", "D1", NoKwargs, NoExisting);

        Assert.True(result.IsValid);
        Assert.Null(result.ErrorMessage);
        var action = Assert.IsType<RibbonAction>(result.Action);
        Assert.Equal("Forecast", action.Name);
        Assert.Equal("model.py", action.Script);
        Assert.Equal("A1:B10", action.Input);
        Assert.Equal("D1", action.Output);
        Assert.Null(action.Kwargs);
    }

    [Fact]
    public void Validate_TrimsAllFields()
    {
        var result = EditActionValidator.Validate(
            "  Forecast  ", "  model.py ", " A1:B10 ", " D1 ", NoKwargs, NoExisting);

        Assert.True(result.IsValid);
        Assert.Equal("Forecast", result.Action!.Name);
        Assert.Equal("model.py", result.Action.Script);
        Assert.Equal("A1:B10", result.Action.Input);
        Assert.Equal("D1", result.Action.Output);
    }

    [Fact]
    public void Validate_WithKwargs_CopiedOntoAction()
    {
        var kwargs = new Dictionary<string, string> { ["periods"] = "12" };
        var result = EditActionValidator.Validate(
            "F", "m.py", "A1", "B1", kwargs, NoExisting);

        Assert.True(result.IsValid);
        Assert.NotNull(result.Action!.Kwargs);
        Assert.Equal("12", result.Action.Kwargs!["periods"]);
        // The action must not alias the caller's dictionary.
        Assert.NotSame(kwargs, result.Action.Kwargs);
    }

    [Fact]
    public void Validate_EmptyKwargs_NormalisedToNull()
    {
        var result = EditActionValidator.Validate(
            "F", "m.py", "A1", "B1", new Dictionary<string, string>(), NoExisting);

        Assert.True(result.IsValid);
        Assert.Null(result.Action!.Kwargs);
    }

    // -------------------------------------------------------------------------
    // Required-field validation
    // -------------------------------------------------------------------------

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_BlankName_Fails(string? name)
    {
        var result = EditActionValidator.Validate(
            name, "m.py", "A1", "B1", NoKwargs, NoExisting);
        Assert.False(result.IsValid);
        Assert.Null(result.Action);
        Assert.Contains("name", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_BlankScript_Fails(string? script)
    {
        var result = EditActionValidator.Validate(
            "F", script, "A1", "B1", NoKwargs, NoExisting);
        Assert.False(result.IsValid);
        Assert.Contains("script", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_BlankInput_Fails(string? input)
    {
        var result = EditActionValidator.Validate(
            "F", "m.py", input, "B1", NoKwargs, NoExisting);
        Assert.False(result.IsValid);
        Assert.Contains("input", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Validate_BlankOutput_Fails(string? output)
    {
        var result = EditActionValidator.Validate(
            "F", "m.py", "A1", output, NoKwargs, NoExisting);
        Assert.False(result.IsValid);
        Assert.Contains("output", result.ErrorMessage!, StringComparison.OrdinalIgnoreCase);
    }

    // -------------------------------------------------------------------------
    // Duplicate-name validation
    // -------------------------------------------------------------------------

    [Fact]
    public void Validate_DuplicateName_AddMode_Fails()
    {
        var existing = new[] { "Forecast", "Clean" };
        var result = EditActionValidator.Validate(
            "Forecast", "m.py", "A1", "B1", NoKwargs, existing, originalName: null);
        Assert.False(result.IsValid);
        Assert.Contains("already exists", result.ErrorMessage!);
    }

    [Fact]
    public void Validate_DuplicateNameAfterTrim_Fails()
    {
        var existing = new[] { "Forecast" };
        var result = EditActionValidator.Validate(
            "  Forecast ", "m.py", "A1", "B1", NoKwargs, existing);
        Assert.False(result.IsValid);
    }

    [Fact]
    public void Validate_EditKeepingSameName_Allowed()
    {
        var existing = new[] { "Forecast", "Clean" };
        var result = EditActionValidator.Validate(
            "Forecast", "m.py", "A1", "B1", NoKwargs, existing, originalName: "Forecast");
        Assert.True(result.IsValid);
    }

    [Fact]
    public void Validate_EditRenameOntoAnotherAction_Fails()
    {
        var existing = new[] { "Forecast", "Clean" };
        var result = EditActionValidator.Validate(
            "Clean", "m.py", "A1", "B1", NoKwargs, existing, originalName: "Forecast");
        Assert.False(result.IsValid);
        Assert.Contains("already exists", result.ErrorMessage!);
    }

    [Fact]
    public void Validate_EditRenameToFreshName_Allowed()
    {
        var existing = new[] { "Forecast", "Clean" };
        var result = EditActionValidator.Validate(
            "Forecast2", "m.py", "A1", "B1", NoKwargs, existing, originalName: "Forecast");
        Assert.True(result.IsValid);
        Assert.Equal("Forecast2", result.Action!.Name);
    }

    [Fact]
    public void Validate_NameCollisionIsCaseSensitive()
    {
        // StateService upserts ordinal/case-sensitive, so "forecast" and
        // "Forecast" are distinct names — not a collision.
        var existing = new[] { "Forecast" };
        var result = EditActionValidator.Validate(
            "forecast", "m.py", "A1", "B1", NoKwargs, existing);
        Assert.True(result.IsValid);
    }

    [Fact]
    public void Validate_NullExistingNames_Throws()
    {
        Assert.Throws<ArgumentNullException>(() =>
            EditActionValidator.Validate("F", "m.py", "A1", "B1", NoKwargs, null!));
    }
}
