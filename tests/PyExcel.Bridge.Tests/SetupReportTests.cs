using PyExcel.Setup;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class SetupReportTests
{
    private static SetupResult Ok(params string[] steps)
    {
        var list = new System.Collections.Generic.List<SetupStep>();
        foreach (var s in steps) list.Add(SetupStep.Ok(s));
        return new SetupResult(list, success: true);
    }

    [Fact]
    public void Headline_Success()
    {
        Assert.Equal("Setup completed successfully.",
            SetupReport.Headline(Ok("probe-python", "provision-venv")));
    }

    [Fact]
    public void Headline_FailureNamesFirstFailedStepAndReason()
    {
        var result = new SetupResult(new[]
        {
            SetupStep.Ok("probe-python"),
            SetupStep.Failed("pip-install", "pip install exited 1: no network"),
        }, success: false);

        Assert.Equal("Setup failed at 'pip-install': pip install exited 1: no network",
            SetupReport.Headline(result));
    }

    [Fact]
    public void Headline_FailureWithoutReason()
    {
        var result = new SetupResult(new[] { SetupStep.Failed("probe-python", "") }, success: false);
        Assert.Equal("Setup failed at 'probe-python'.", SetupReport.Headline(result));
    }

    [Fact]
    public void Summarize_ListsEveryStepThenHeadline()
    {
        var result = new SetupResult(new[]
        {
            SetupStep.Ok("probe-python"),
            SetupStep.Failed("provision-venv", "access denied"),
        }, success: false);

        var text = SetupReport.Summarize(result);

        Assert.Contains("[ok]   probe-python", text);
        Assert.Contains("[fail] provision-venv — access denied", text);
        Assert.Contains("Setup failed at 'provision-venv': access denied", text);
    }
}
