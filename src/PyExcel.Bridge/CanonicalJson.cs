using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace PyExcel.Bridge;

/// <summary>
/// Minimal canonical JSON encoder/decoder used for frame meta.
///
/// Mirrors what Python's <c>json.dumps(meta, ensure_ascii=False,
/// separators=(",", ":"), sort_keys=True)</c> produces, so the on-wire
/// bytes are deterministic across both sides of the protocol. Pure-stdlib
/// by design — no Newtonsoft.Json, no System.Text.Json. The .xll bundle
/// already has enough surface; framing must stay loadable without
/// dragging in JSON-library version conflicts.
///
/// Supported value types:
/// <list type="bullet">
///   <item><c>null</c></item>
///   <item><see cref="bool"/></item>
///   <item>Integer types (<see cref="sbyte"/>, <see cref="byte"/>,
///         <see cref="short"/>, <see cref="ushort"/>, <see cref="int"/>,
///         <see cref="uint"/>, <see cref="long"/>) — emitted without
///         decimal point</item>
///   <item><see cref="float"/>, <see cref="double"/> — round-trip
///         format (<c>R</c>), matches Python for typical values</item>
///   <item><see cref="string"/> — UTF-8 with Python-compatible escaping</item>
///   <item><see cref="IDictionary{TKey,TValue}"/> keyed by string — keys
///         emitted in ordinal sort order</item>
///   <item><see cref="IEnumerable"/> of supported values — preserved order</item>
/// </list>
///
/// Decoded numbers are <see cref="long"/> if integer-valued and fit in 64
/// bits, otherwise <see cref="double"/>. Decoded objects are
/// <see cref="Dictionary{TKey,TValue}"/>, decoded arrays are
/// <see cref="List{T}"/>.
/// </summary>
internal static class CanonicalJson
{
    // -------------------------------------------------------------------------
    // Encode
    // -------------------------------------------------------------------------

    public static byte[] Encode(object? value)
    {
        var sb = new StringBuilder();
        EncodeValue(value, sb);
        return Encoding.UTF8.GetBytes(sb.ToString());
    }

    private static void EncodeValue(object? value, StringBuilder sb)
    {
        if (value is null)
        {
            sb.Append("null");
            return;
        }
        switch (value)
        {
            case bool b:
                sb.Append(b ? "true" : "false");
                return;
            case string s:
                EncodeString(s, sb);
                return;
            case sbyte or byte or short or ushort or int or uint or long:
                sb.Append(((IConvertible)value).ToInt64(CultureInfo.InvariantCulture)
                    .ToString(CultureInfo.InvariantCulture));
                return;
            case ulong u:
                if (u > long.MaxValue)
                    throw new ArgumentException($"ulong {u} exceeds JSON-safe range");
                sb.Append(((long)u).ToString(CultureInfo.InvariantCulture));
                return;
            case float f:
                EncodeDouble(f, sb);
                return;
            case double d:
                EncodeDouble(d, sb);
                return;
            case IDictionary dict:
                EncodeObject(dict, sb);
                return;
            case IEnumerable seq:
                EncodeArray(seq, sb);
                return;
            default:
                throw new ArgumentException(
                    $"value of type {value.GetType().FullName} is not JSON-encodable");
        }
    }

    private static void EncodeDouble(double d, StringBuilder sb)
    {
        if (double.IsNaN(d) || double.IsInfinity(d))
            throw new ArgumentException($"non-finite number {d} is not JSON-encodable");

        // Match Python json: integer-valued floats keep ".0" (e.g. 1.0 → "1.0").
        // "R" gives shortest round-trippable repr on .NET Core 3.0+.
        var s = d.ToString("R", CultureInfo.InvariantCulture);
        if (s.IndexOf('.') < 0 && s.IndexOf('e') < 0 && s.IndexOf('E') < 0 && s.IndexOf('n') < 0)
            s += ".0";
        sb.Append(s);
    }

    private static void EncodeString(string s, StringBuilder sb)
    {
        sb.Append('"');
        foreach (var c in s)
        {
            switch (c)
            {
                case '\\': sb.Append("\\\\"); break;
                case '"': sb.Append("\\\""); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                default:
                    if (c < 0x20)
                        sb.AppendFormat(CultureInfo.InvariantCulture, "\\u{0:x4}", (int)c);
                    else
                        sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
    }

    private static void EncodeObject(IDictionary dict, StringBuilder sb)
    {
        var keys = new List<string>(dict.Count);
        foreach (var k in dict.Keys)
        {
            if (k is not string ks)
                throw new ArgumentException(
                    $"object keys must be strings, got {k?.GetType().FullName ?? "null"}");
            keys.Add(ks);
        }
        keys.Sort(StringComparer.Ordinal);

        sb.Append('{');
        for (var i = 0; i < keys.Count; i++)
        {
            if (i > 0) sb.Append(',');
            EncodeString(keys[i], sb);
            sb.Append(':');
            EncodeValue(dict[keys[i]], sb);
        }
        sb.Append('}');
    }

    private static void EncodeArray(IEnumerable seq, StringBuilder sb)
    {
        sb.Append('[');
        var first = true;
        foreach (var item in seq)
        {
            if (!first) sb.Append(',');
            first = false;
            EncodeValue(item, sb);
        }
        sb.Append(']');
    }

    // -------------------------------------------------------------------------
    // Decode
    // -------------------------------------------------------------------------

    // Strict UTF-8 — throw on invalid byte sequences instead of substituting
    // U+FFFD, matching the behaviour of Python's bytes.decode("utf-8").
    private static readonly UTF8Encoding StrictUtf8 = new(
        encoderShouldEmitUTF8Identifier: false,
        throwOnInvalidBytes: true);

    public static object? Decode(byte[] utf8)
    {
        if (utf8 is null) throw new ArgumentNullException(nameof(utf8));
        string text;
        try
        {
            text = StrictUtf8.GetString(utf8);
        }
        catch (DecoderFallbackException exc)
        {
            throw new FormatException("invalid UTF-8 byte sequence", exc);
        }
        var parser = new Parser(text);
        var value = parser.ParseValue();
        parser.SkipWhitespace();
        if (!parser.AtEnd)
            throw new FormatException($"trailing data at position {parser.Position}");
        return value;
    }

    private struct Parser
    {
        private readonly string _text;
        private int _pos;

        public Parser(string text)
        {
            _text = text;
            _pos = 0;
        }

        public int Position => _pos;
        public bool AtEnd => _pos >= _text.Length;

        public void SkipWhitespace()
        {
            while (_pos < _text.Length)
            {
                var c = _text[_pos];
                if (c == ' ' || c == '\t' || c == '\n' || c == '\r')
                    _pos++;
                else
                    break;
            }
        }

        public object? ParseValue()
        {
            SkipWhitespace();
            if (AtEnd) throw new FormatException("unexpected end of input");

            var c = _text[_pos];
            return c switch
            {
                '{' => ParseObject(),
                '[' => ParseArray(),
                '"' => ParseString(),
                't' or 'f' => ParseBool(),
                'n' => ParseNull(),
                _ => ParseNumber(),
            };
        }

        private Dictionary<string, object?> ParseObject()
        {
            Expect('{');
            var result = new Dictionary<string, object?>(StringComparer.Ordinal);
            SkipWhitespace();
            if (Peek() == '}') { _pos++; return result; }
            while (true)
            {
                SkipWhitespace();
                var key = ParseString();
                SkipWhitespace();
                Expect(':');
                var value = ParseValue();
                result[key] = value;
                SkipWhitespace();
                var c = NextOrThrow();
                if (c == ',') continue;
                if (c == '}') return result;
                throw new FormatException($"expected ',' or '}}' at position {_pos - 1}, got '{c}'");
            }
        }

        private List<object?> ParseArray()
        {
            Expect('[');
            var result = new List<object?>();
            SkipWhitespace();
            if (Peek() == ']') { _pos++; return result; }
            while (true)
            {
                var value = ParseValue();
                result.Add(value);
                SkipWhitespace();
                var c = NextOrThrow();
                if (c == ',') continue;
                if (c == ']') return result;
                throw new FormatException($"expected ',' or ']' at position {_pos - 1}, got '{c}'");
            }
        }

        private string ParseString()
        {
            Expect('"');
            var sb = new StringBuilder();
            while (true)
            {
                if (_pos >= _text.Length)
                    throw new FormatException("unterminated string");
                var c = _text[_pos++];
                if (c == '"') return sb.ToString();
                if (c == '\\')
                {
                    if (_pos >= _text.Length)
                        throw new FormatException("unterminated escape");
                    var esc = _text[_pos++];
                    switch (esc)
                    {
                        case '"': sb.Append('"'); break;
                        case '\\': sb.Append('\\'); break;
                        case '/': sb.Append('/'); break;
                        case 'b': sb.Append('\b'); break;
                        case 'f': sb.Append('\f'); break;
                        case 'n': sb.Append('\n'); break;
                        case 'r': sb.Append('\r'); break;
                        case 't': sb.Append('\t'); break;
                        case 'u':
                            if (_pos + 4 > _text.Length)
                                throw new FormatException("truncated \\u escape");
                            var hex = _text.Substring(_pos, 4);
                            _pos += 4;
                            if (!ushort.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code))
                                throw new FormatException($"invalid \\u escape: {hex}");
                            sb.Append((char)code);
                            break;
                        default:
                            throw new FormatException($"invalid escape \\{esc}");
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }
        }

        private bool ParseBool()
        {
            if (TryMatch("true")) return true;
            if (TryMatch("false")) return false;
            throw new FormatException($"invalid literal at position {_pos}");
        }

        private object? ParseNull()
        {
            if (TryMatch("null")) return null;
            throw new FormatException($"invalid literal at position {_pos}");
        }

        private object ParseNumber()
        {
            var start = _pos;
            if (Peek() == '-') _pos++;
            while (_pos < _text.Length)
            {
                var c = _text[_pos];
                if ((c >= '0' && c <= '9') || c == '.' || c == 'e' || c == 'E' || c == '+' || c == '-')
                    _pos++;
                else
                    break;
            }
            var slice = _text.Substring(start, _pos - start);
            if (slice.Length == 0)
                throw new FormatException($"expected number at position {start}");

            // Integer if no fractional/exponent part and it fits in long.
            if (slice.IndexOf('.') < 0 && slice.IndexOf('e') < 0 && slice.IndexOf('E') < 0)
            {
                if (long.TryParse(slice, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l))
                    return l;
            }
            if (!double.TryParse(slice, NumberStyles.Float, CultureInfo.InvariantCulture, out var d))
                throw new FormatException($"invalid number '{slice}'");
            return d;
        }

        private char Peek() => _pos < _text.Length ? _text[_pos] : '\0';

        private char NextOrThrow()
        {
            if (_pos >= _text.Length) throw new FormatException("unexpected end of input");
            return _text[_pos++];
        }

        private void Expect(char c)
        {
            if (_pos >= _text.Length || _text[_pos] != c)
                throw new FormatException(
                    $"expected '{c}' at position {_pos}, got '{(_pos < _text.Length ? _text[_pos] : ' ')}'");
            _pos++;
        }

        private bool TryMatch(string literal)
        {
            if (_pos + literal.Length > _text.Length) return false;
            for (var i = 0; i < literal.Length; i++)
                if (_text[_pos + i] != literal[i])
                    return false;
            _pos += literal.Length;
            return true;
        }
    }
}
