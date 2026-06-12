using System;
using System.IO;
using System.Text;
using PyExcel.Common.Logging;

namespace PyExcel.Setup;

/// <summary>
/// Prepares the user-facing parts of a PyExcel project directory: the
/// <c>userScripts</c> folder where transform scripts live, seeded with a
/// runnable example the first time. This complements the <c>.pyexcel-venv</c>
/// and <c>.pyexcel-kernel</c> that the rest of <see cref="SetupService"/>
/// provisions — those are the machinery; <c>userScripts</c> is what the user
/// actually edits. Cross-platform and unit-tested.
/// </summary>
public sealed class ProjectScaffolder
{
    /// <summary>Folder under the project directory that holds the user's
    /// transform scripts. Must match the ribbon's <c>UserScriptsDir()</c> and
    /// <c>ScriptDirectoryWatcher</c> so Setup writes where the runtime reads.</summary>
    public const string UserScriptsDirName = "userScripts";

    /// <summary>Name of the starter script dropped into an empty project.</summary>
    public const string ExampleScriptName = "example.py";

    private static readonly string ExampleScript =
        "# PyExcel example script.\n" +
        "#\n" +
        "# A PyExcel action runs transform(inputs) and writes the returned dict\n" +
        "# of named results back to the sheet. Copy or edit this file, then pick\n" +
        "# it in the Action dialog.\n" +
        "\n" +
        "from typing import Any, Dict\n" +
        "\n" +
        "\n" +
        "def transform(inputs: Dict[str, Any]) -> Dict[str, Any]:\n" +
        "    # inputs maps each named input range to a value:\n" +
        "    #   multi-row/column range -> pandas.DataFrame\n" +
        "    #   single row or column   -> list\n" +
        "    #   single cell            -> scalar (int/float/bool/str/Timestamp)\n" +
        "    #\n" +
        "    # This example sums every numeric input cell and returns the total.\n" +
        "    total = 0.0\n" +
        "    for value in inputs.values():\n" +
        "        try:\n" +
        "            total += float(value)\n" +
        "        except (TypeError, ValueError):\n" +
        "            pass\n" +
        "    return {\"total\": total}\n";

    private readonly ILog _log;

    public ProjectScaffolder(ILog? log = null) => _log = log ?? NullLog.Instance;

    /// <summary>
    /// Create <c><paramref name="projectDir"/>/userScripts</c> (if missing) and,
    /// when it holds no scripts yet, drop a starter <c>example.py</c>. Never
    /// overwrites an existing script. Returns the userScripts path.
    /// </summary>
    public string Scaffold(string projectDir)
    {
        if (string.IsNullOrWhiteSpace(projectDir))
            throw new ArgumentException("project directory required", nameof(projectDir));

        var userScripts = Path.Combine(projectDir, UserScriptsDirName);
        Directory.CreateDirectory(userScripts);

        if (HasAnyScript(userScripts))
        {
            _log.Info($"scaffold: {UserScriptsDirName}/ already has scripts; left as-is");
        }
        else
        {
            File.WriteAllText(
                Path.Combine(userScripts, ExampleScriptName), ExampleScript, new UTF8Encoding(false));
            _log.Info($"scaffold: wrote starter {UserScriptsDirName}/{ExampleScriptName}");
        }

        return userScripts;
    }

    private static bool HasAnyScript(string dir)
    {
        try { return Directory.GetFiles(dir, "*.py").Length > 0; }
        catch { return false; }
    }
}
