using System;
using System.Text;

namespace PyExcel.State;

/// <summary>
/// Decodes the <c>RefersTo</c> formula of a v1 PyExcel defined Name back into
/// the string value it carries.
///
/// <para>v1 stored every value as an Excel string-literal formula. Short
/// values are a single literal — <c>="A1:C10"</c> — with any embedded quote
/// doubled (<c>"</c> → <c>""</c>). Values over Excel's 255-char literal cap
/// (notably the serialized <c>Actions</c> list) were written as a chain of
/// literals joined by <c>&amp;</c> — <c>="chunk1"&amp;"chunk2"&amp;…</c> — which
/// Excel evaluates back to the concatenation. This decoder is the inverse of
/// both forms, so the COM-side reader can recover the value from
/// <c>Name.RefersTo</c> without calling <c>Evaluate</c> (which would force a
/// recalc and depends on the name's scope being active).</para>
///
/// <para>A formula that isn't a string-literal chain (e.g. a Name that refers
/// to a real range, or anything PyExcel never wrote) decodes to
/// <see langword="null"/> — the reader treats that as "not a PyExcel
/// value".</para>
/// </summary>
public static class LegacyFormulaDecoder
{
    /// <summary>Decode a <c>RefersTo</c> formula into its string value, or
    /// <see langword="null"/> if <paramref name="refersTo"/> is blank or isn't
    /// a well-formed string-literal chain.</summary>
    public static string? Decode(string? refersTo)
    {
        if (string.IsNullOrEmpty(refersTo)) return null;

        var s = refersTo!.Trim();
        if (s.Length == 0) return null;
        if (s[0] == '=') s = s.Substring(1).Trim();

        var result = new StringBuilder();
        int i = 0;
        bool any = false;

        while (i < s.Length)
        {
            // Each segment must be a quoted string literal.
            if (s[i] != '"') return null;
            i++; // consume opening quote

            while (true)
            {
                if (i >= s.Length) return null; // unterminated literal
                char c = s[i];
                if (c == '"')
                {
                    // A doubled quote ("") is a literal quote; a lone quote
                    // closes the literal.
                    if (i + 1 < s.Length && s[i + 1] == '"')
                    {
                        result.Append('"');
                        i += 2;
                        continue;
                    }
                    i++; // consume closing quote
                    break;
                }
                result.Append(c);
                i++;
            }

            any = true;

            // After a literal: either the formula ends, or a '&' joins the
            // next literal. Whitespace around '&' is tolerated.
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i >= s.Length) break;
            if (s[i] != '&') return null;
            i++; // consume '&'
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
        }

        return any ? result.ToString() : null;
    }
}
