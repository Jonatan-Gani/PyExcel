using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// Parses the ribbon's Input / Output text fields into an ordered list of
/// named (or anonymous) range bindings.
///
/// <para>Syntax — semicolon-separated bindings, each one of:</para>
/// <list type="bullet">
///   <item><c>A1:C10</c> — anonymous, untyped range. The kernel names it
///     by its resolved type and ordinal (<c>df1</c>, <c>list1</c>, …) and
///     builds whatever the range's dimensions imply.</item>
///   <item><c>prices=A1:C10</c> — named range. The name is the key the
///     range appears under in the <c>inputs</c> dict handed to
///     <c>transform</c>.</item>
///   <item><c>prices:dataframe=A1:C10</c> — named range with a declared
///     type. The kernel constructs exactly that type or fails with a
///     message naming the binding.</item>
///   <item><c>:list=A1:A10</c> — anonymous range with a declared type.</item>
/// </list>
///
/// <para>The type token is recognised only when the text after the final
/// <c>:</c> of the name segment parses as a known <see cref="PyExcelType"/>.
/// A name that merely happens to contain a colon keeps its colon and is
/// read as a name, so existing saved bindings cannot be silently
/// reinterpreted as typed ones.</para>
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
            var declaredType = PyExcelType.Auto;

            var eq = entry.IndexOf('=');
            if (eq < 0)
            {
                name = null;
                rangeText = entry;
            }
            else
            {
                // Only the segment left of the FIRST '=' is name-and-type;
                // everything right of it is range text handed verbatim to
                // Excel, so a declared type must never live there.
                var head = entry.Substring(0, eq).Trim();
                rangeText = entry.Substring(eq + 1).Trim();

                name = SplitDeclaredType(head, out declaredType);

                if (name.Length == 0)
                {
                    // ':list=A1:A10' is a legitimate anonymous typed
                    // binding; a bare '=A1:A10' is still a typo.
                    if (declaredType == PyExcelType.Auto)
                        throw new FormatException(
                            $"empty name before '=' in '{rawEntry}'");
                    name = null;
                }
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

            result.Add(new RangeBinding(name, rangeText, declaredType));
        }

        return result;
    }

    /// <summary>
    /// Split a binding's name-and-type segment (everything left of the
    /// <c>=</c>) into its name and declared type.
    ///
    /// <para>The split happens at the LAST <c>:</c>, and only when the text
    /// after it parses as a known type. That rule is what keeps the
    /// extension backward compatible: a previously saved name containing a
    /// colon does not parse as a type, so it survives intact as a name.</para>
    /// </summary>
    /// <returns>The name segment, trimmed; empty when the binding is
    /// anonymous.</returns>
    private static string SplitDeclaredType(string head, out PyExcelType declaredType)
    {
        declaredType = PyExcelType.Auto;

        var colon = head.LastIndexOf(':');
        if (colon < 0) return head;

        var suffix = head.Substring(colon + 1);
        if (!PyExcelTypes.TryParse(suffix, out var parsed)) return head;

        declaredType = parsed;
        return head.Substring(0, colon).Trim();
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
            var named = name is not null && name.Length > 0;
            var typed = b.DeclaredType != PyExcelType.Auto;

            if (named) sb.Append(name);
            if (typed)
            {
                sb.Append(':');
                sb.Append(PyExcelTypes.WireName(b.DeclaredType));
            }
            if (named || typed) sb.Append('=');
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
/// <see cref="DeclaredType"/> is the type the user picked in the action
/// dialog's type box; <see cref="PyExcelType.Auto"/> (the default) defers
/// to the range's measured dimensions at run time.
/// </summary>
public sealed record RangeBinding(
    string? Name,
    string RangeText,
    PyExcelType DeclaredType = PyExcelType.Auto);
