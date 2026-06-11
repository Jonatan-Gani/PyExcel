using System;
using System.IO;
using System.Text;

namespace PyExcel.Forms;

/// <summary>
/// Creates a new user script from a starter template (Note 2). Cross-platform
/// and unit-tested — the WinForms EditAction dialog calls <see cref="Create"/>
/// when the user clicks "New…", so first use isn't blocked by an empty script
/// list. Kept out of <c>#if NETFRAMEWORK</c> so the logic runs on CI.
/// </summary>
public static class ScriptScaffold
{
    /// <summary>
    /// The starter <c>transform()</c> skeleton written into a new script.
    /// Mirrors the authoring contract documented in the README. Uses
    /// <c>#</c> comments (not a triple-quoted docstring) to keep the C#
    /// literal readable, and LF line endings (Python is happy with LF on
    /// Windows).
    /// </summary>
    public const string Template =
        "# PyExcel user script.\n" +
        "#\n" +
        "# Define transform(inputs) and return a dict of named results.\n" +
        "# See the README for the full input/output contract.\n" +
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
        "    # Return a dict of result names to values; each is written back by\n" +
        "    # type (DataFrame -> table, list -> spill range, scalar -> single\n" +
        "    # cell, Plotly/Matplotlib figure -> chart/image).\n" +
        "    first = next(iter(inputs.values()), None)\n" +
        "    return {\"result\": first}\n";

    /// <summary>
    /// Turn user input into a safe <c>&lt;name&gt;.py</c> file name: keep
    /// letters, digits, <c>_</c> and <c>-</c>, replace everything else with
    /// <c>_</c>, drop any directory or existing <c>.py</c> the user typed, and
    /// re-add <c>.py</c>. Throws if nothing usable remains.
    /// </summary>
    public static string SanitizeFileName(string? desiredName)
    {
        var name = (desiredName ?? string.Empty).Trim();
        if (name.Length == 0)
            throw new ArgumentException("Enter a script name.", nameof(desiredName));

        // Strip a trailing ".py" (any case) so it isn't doubled.
        if (name.Length > 3 && name.EndsWith(".py", StringComparison.OrdinalIgnoreCase))
            name = name.Substring(0, name.Length - 3);

        var sb = new StringBuilder(name.Length);
        foreach (var ch in name)
        {
            bool safe = (ch >= 'a' && ch <= 'z')
                || (ch >= 'A' && ch <= 'Z')
                || (ch >= '0' && ch <= '9')
                || ch == '_' || ch == '-';
            sb.Append(safe ? ch : '_');
        }

        var stem = sb.ToString().Trim('_', '-');
        if (stem.Length == 0)
            throw new ArgumentException(
                "That script name has no usable characters.", nameof(desiredName));

        return stem + ".py";
    }

    /// <summary>
    /// Create a new script under <paramref name="userScriptsDir"/> from
    /// <see cref="Template"/> and return its file name (not the full path).
    /// The directory is created if missing; name collisions are resolved by
    /// appending <c>_1</c>, <c>_2</c>, … so an existing script is never
    /// overwritten.
    /// </summary>
    public static string Create(string userScriptsDir, string? desiredName)
    {
        if (string.IsNullOrWhiteSpace(userScriptsDir))
            throw new ArgumentException("userScriptsDir is required.", nameof(userScriptsDir));

        var fileName = SanitizeFileName(desiredName);
        Directory.CreateDirectory(userScriptsDir);

        var stem = fileName.Substring(0, fileName.Length - 3); // drop ".py"
        var candidate = fileName;
        var n = 1;
        while (File.Exists(Path.Combine(userScriptsDir, candidate)))
        {
            candidate = $"{stem}_{n}.py";
            n++;
        }

        File.WriteAllText(
            Path.Combine(userScriptsDir, candidate), Template, new UTF8Encoding(false));
        return candidate;
    }
}
