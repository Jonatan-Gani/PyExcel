using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// Parses the ribbon's Input / Output text fields into an ordered list of
/// named (or anonymous) range bindings.
///
/// <para>Syntax — semicolon-separated bindings, each one of:</para>
/// <list type="bullet">
///   <item><c>A1:C10</c> — anonymous range, picked up as the next
///     positional argument to the user's transform function.</item>
///   <item><c>prices=A1:C10</c> — named range. The dispatcher passes it
///     as a positional argument in declaration order; the name is the
///     hint the user assigned to their <c>transform(prices, …)</c>
///     parameter and is preserved on the binding for future Phase-8 UI
///     work.</item>
/// </list>
///
/// <para>Whitespace around the separator, the <c>=</c>, and the range
/// text itself is trimmed. Empty entries (a stray <c>;;</c> or trailing
/// <c>;</c>) are silently skipped — they're a common typing accident,
/// not a usage error.</para>
///
/// <para>Returns an empty list for <see langword="null"/>, empty, or
/// whitespace-only input; the dispatcher treats that as "no positional
/// arguments". A missing range text after an <c>=</c> (e.g.
/// <c>prices=</c>) or a missing name before it (e.g. <c>=A1:C10</c>) is
/// a real user error and throws <see cref="FormatException"/>.</para>
///
/// <para>Pure logic — no Excel COM, no range validation. The dispatcher
/// is responsible for resolving each <see cref="RangeBinding.RangeText"/>
/// into a real <c>Range</c> via the host application.</para>
/// </summary>
public static class RibbonRangeParser
{
    /// <summary>Parse the ribbon-input text into ordered bindings.</summary>
    /// <exception cref="FormatException">An entry is malformed — empty
    /// name, empty range, or duplicate name.</exception>
    public static IReadOnlyList<RangeBinding> Parse(string? input)
    {
        if (string.IsNullOrWhiteSpace(input)) return Array.Empty<RangeBinding>();

        var result = new List<RangeBinding>();
        HashSet<string>? seenNames = null;

        foreach (var rawEntry in input!.Split(';'))
        {
            var entry = rawEntry.Trim();
            if (entry.Length == 0) continue;

            string? name;
            string rangeText;

            var eq = entry.IndexOf('=');
            if (eq < 0)
            {
                name = null;
                rangeText = entry;
            }
            else
            {
                name = entry.Substring(0, eq).Trim();
                rangeText = entry.Substring(eq + 1).Trim();
                if (name.Length == 0)
                    throw new FormatException(
                        $"empty name before '=' in '{rawEntry}'");
            }

            if (rangeText.Length == 0)
                throw new FormatException(
                    $"empty range text in '{rawEntry}'");

            if (name is not null)
            {
                seenNames ??= new HashSet<string>(StringComparer.Ordinal);
                if (!seenNames.Add(name))
                    throw new FormatException(
                        $"duplicate name '{name}' in input");
            }

            result.Add(new RangeBinding(name, rangeText));
        }

        return result;
    }

    /// <summary>
    /// Serialise bindings back into the ribbon text syntax that
    /// <see cref="Parse"/> reads — the inverse of <see cref="Parse"/>. A named
    /// binding becomes <c>name=range</c>; an anonymous one is just the range;
    /// entries join with <c>"; "</c>. Rows whose range text is blank are
    /// skipped. <c>Parse(Format(x))</c> round-trips back to the same bindings.
    /// </summary>
    public static string Format(IEnumerable<RangeBinding> bindings)
    {
        if (bindings is null) throw new ArgumentNullException(nameof(bindings));

        var sb = new System.Text.StringBuilder();
        foreach (var b in bindings)
        {
            if (b is null) continue;
            var range = (b.RangeText ?? string.Empty).Trim();
            if (range.Length == 0) continue;

            if (sb.Length > 0) sb.Append("; ");

            var name = b.Name?.Trim();
            if (name is not null && name.Length > 0)
            {
                sb.Append(name);
                sb.Append('=');
            }
            sb.Append(range);
        }
        return sb.ToString();
    }
}

/// <summary>
/// One parsed binding from the ribbon's Input or Output text.
/// <see cref="Name"/> is <see langword="null"/> for an anonymous entry;
/// <see cref="RangeText"/> is the unresolved range text (e.g.
/// <c>"Sheet1!A1:C10"</c>) — the dispatcher resolves it via Excel COM.
/// </summary>
public sealed record RangeBinding(string? Name, string RangeText);
