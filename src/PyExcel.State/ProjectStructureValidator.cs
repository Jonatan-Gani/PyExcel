using System.Collections.Generic;
using System.IO;

namespace PyExcel.State;

/// <summary>
/// The outcome of validating that an enabled workbook's on-disk project structure
/// is present: the Python virtual environment, the extracted kernel, and the
/// userScripts folder that <c>PyExcel.Setup</c> provisions on Enable.
/// </summary>
public sealed record ProjectStructureCheck(bool Ok, IReadOnlyList<string> Missing)
{
    /// <summary>The all-present result.</summary>
    public static readonly ProjectStructureCheck Healthy =
        new(true, System.Array.Empty<string>());
}

/// <summary>
/// Fast, file-only validation that the structure <c>PyExcel.Setup</c> creates for
/// an enabled workbook still exists in its project directory. Used by the open
/// hook to tell the user up front — before they click Run and hit a cryptic kernel
/// error — when the environment is missing (e.g. the project folder was copied
/// without <c>.pyexcel-venv</c>, or the venv was deleted).
///
/// <para>It deliberately does <b>not</b> spawn Python or import anything — it's a
/// handful of <see cref="File.Exists"/> / <see cref="Directory.Exists"/> checks, so
/// it's safe to run synchronously on a workbook-open event. A full dependency
/// import check is the heavier job that belongs to Setup's
/// <c>DependencyVerifier</c>.</para>
/// </summary>
public static class ProjectStructureValidator
{
    /// <summary>The per-project venv directory Setup provisions.</summary>
    public const string VenvDirName = ".pyexcel-venv";

    /// <summary>The directory Setup extracts the kernel package into.</summary>
    public const string KernelDirName = ".pyexcel-kernel";

    /// <summary>The user scripts folder Setup scaffolds.</summary>
    public const string ScriptsDirName = "userScripts";

    /// <summary>Validate the structure under <paramref name="projectDir"/>. A
    /// null/blank/missing directory, or any missing component, yields
    /// <c>Ok == false</c> with a human-readable list of what's missing.</summary>
    public static ProjectStructureCheck Validate(string? projectDir)
    {
        if (string.IsNullOrWhiteSpace(projectDir) || !Directory.Exists(projectDir))
            return new ProjectStructureCheck(false, new[] { "the project folder" });

        var missing = new List<string>();

        // The venv interpreter — check both layouts so the result doesn't depend
        // on which OS provisioned the project (Windows: Scripts\python.exe; POSIX:
        // bin/python).
        var venv = Path.Combine(projectDir!, VenvDirName);
        var hasVenv = File.Exists(Path.Combine(venv, "Scripts", "python.exe"))
                      || File.Exists(Path.Combine(venv, "bin", "python"));
        if (!hasVenv) missing.Add("the Python virtual environment (.pyexcel-venv)");

        // The extracted kernel package, keyed on the canonical importable marker.
        var kernelMain = Path.Combine(projectDir!, KernelDirName, "pyexcel", "kernel", "__main__.py");
        if (!File.Exists(kernelMain)) missing.Add("the PyExcel kernel (.pyexcel-kernel)");

        // The user scripts folder.
        if (!Directory.Exists(Path.Combine(projectDir!, ScriptsDirName)))
            missing.Add("the userScripts folder");

        return missing.Count == 0
            ? ProjectStructureCheck.Healthy
            : new ProjectStructureCheck(false, missing);
    }
}
