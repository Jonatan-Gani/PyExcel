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

    /// <summary>Name of the project README the ribbon's "Read Me" button opens.</summary>
    public const string ReadmeName = "README.md";

    /// <summary>Default README content written into a new project so the
    /// "Read Me" ribbon button always has something to open.</summary>
    public static readonly string ReadmeContent =
        "# PyExcel project\n" +
        "\n" +
        "This folder holds the PyExcel environment for your workbook:\n" +
        "\n" +
        "- `userScripts/` — your Python transform scripts (start from `example.py`).\n" +
        "- `.pyexcel-venv/` — the private Python environment PyExcel created.\n" +
        "- `.pyexcel-kernel/` — the PyExcel kernel that runs your scripts.\n" +
        "\n" +
        "## Using PyExcel\n" +
        "\n" +
        "1. **Edit** a script: pick it in the ribbon's *Script* box and click *Edit*.\n" +
        "2. **Add an action**: click *Add* to bind a script to input/output ranges.\n" +
        "3. **Run**: select the action and click *Run*; output is written to the\n" +
        "   range you configured, and `print()` output shows in the log window.\n" +
        "4. **Errors**: a failed run shows the Python traceback; use *Show Last\n" +
        "   Error* to see it again.\n" +
        "\n" +
        "Your actions and settings are saved inside the workbook, so they travel\n" +
        "with the file.\n";

    private static readonly string ExampleScript =
        "# PyExcel example script.\n" +
        "#\n" +
        "# A PyExcel action runs transform(inputs) and writes the returned dict\n" +
        "# of named results back to the sheet. Copy or edit this file, then pick\n" +
        "# it in the Action dialog.\n" +
        "#\n" +
        "# Note: scripts run in a kernel with no console — input() and reading\n" +
        "# sys.stdin are disabled and raise an error. Read values from 'inputs'.\n" +
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

        // Drop a README at the project root the "Read Me" ribbon button opens.
        // Never overwrite a user-edited one.
        var readme = Path.Combine(projectDir, ReadmeName);
        if (!File.Exists(readme))
        {
            File.WriteAllText(readme, ReadmeContent, new UTF8Encoding(false));
            _log.Info($"scaffold: wrote {ReadmeName}");
        }

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
