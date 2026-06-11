using System;
using System.Collections.Generic;
using System.Text;

namespace PyExcel.Forms;

/// <summary>
/// Round-trips a <see cref="RibbonAction"/>'s optional keyword arguments
/// between the in-memory dictionary and the one-pair-per-line
/// <c>key=value</c> text the EditAction form edits them as.
///
/// <para>Pure logic, no WinForms — so the parse/format rules the dialog
/// relies on are unit-testable on Linux CI. The form is just a multiline
/// text box around this.</para>
/// </summary>
public static class KwargsText
{
    /// <summary>
    /// Parse the multiline <c>key=value</c> text into an ordered
    /// dictionary. Blank lines are ignored. The split is on the FIRST
    /// <c>=</c> so a value may itself contain <c>=</c>
    /// (<c>expr=a==b</c> → key <c>expr</c>, value <c>a==b</c>). Keys are
    /// trimmed; values keep their interior whitespace but are trimmed at
    /// the ends.
    /// </summary>
    /// <param name="text">The raw text-box contents (may be null).</param>
    /// <param name="error">On a malformed line, the reason; else null.</param>
    /// <returns>The parsed pairs, or null when <paramref name="error"/> is
    /// set. An empty/whitespace input parses to an empty dictionary, not
    /// an error.</returns>
    public static IReadOnlyDictionary<string, string>? TryParse(string? text, out string? error)
    {
        error = null;
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(text)) return result;

        // Tolerate all three line endings the way the rest of the codebase
        // does — normalise CRLF / CR to LF before splitting.
        var normalised = text!.Replace("\r\n", "\n").Replace('\r', '\n');
        foreach (var rawLine in normalised.Split('\n'))
        {
            var line = rawLine.Trim();
            if (line.Length == 0) continue;

            var eq = line.IndexOf('=');
            if (eq < 0)
            {
                error = $"Keyword argument '{line}' is missing an '=' " +
                        "(each line must be name=value).";
                return null;
            }

            var key = line.Substring(0, eq).Trim();
            if (key.Length == 0)
            {
                error = "A keyword-argument line has a blank name before '='.";
                return null;
            }
            if (result.ContainsKey(key))
            {
                error = $"Keyword argument '{key}' is listed more than once.";
                return null;
            }

            var value = line.Substring(eq + 1).Trim();
            result[key] = value;
        }

        return result;
    }

    /// <summary>
    /// Render a kwargs dictionary back to the one-pair-per-line text the
    /// form edits — the inverse of <see cref="TryParse"/>. A null or empty
    /// dictionary renders as the empty string.
    /// </summary>
    public static string Format(IReadOnlyDictionary<string, string>? kwargs)
    {
        if (kwargs is null || kwargs.Count == 0) return string.Empty;
        var sb = new StringBuilder();
        var first = true;
        foreach (var kv in kwargs)
        {
            if (!first) sb.Append('\n');
            sb.Append(kv.Key).Append('=').Append(kv.Value);
            first = false;
        }
        return sb.ToString();
    }
}
