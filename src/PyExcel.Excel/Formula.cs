using System;

namespace PyExcel.Excel;

/// <summary>
/// An Excel A1-mode formula (e.g. <c>=SUM(A1:B2)</c>). The kernel returns
/// instances of this type when a user <c>transform()</c> emits a live
/// formula instead of a pre-computed value; the host writes them via
/// <c>Range.Formula</c> rather than <c>Range.Value2</c> so Excel recomputes
/// them on every recalc.
///
/// <para>The wire representation is an Arrow string column whose field
/// carries a <c>pyexcel-cell-type=formula</c> metadata key. See
/// <see cref="ArrowMarshal"/> for the encode/decode plumbing and
/// <c>embedded/pyexcel/kernel/types.py</c> for the matching Python
/// dataclass.</para>
/// </summary>
public sealed class Formula : IEquatable<Formula>
{
    /// <summary>The formula source, always starting with <c>=</c>.</summary>
    public string Text { get; }

    /// <param name="text">The formula source. Must start with <c>=</c> —
    /// anything else throws <see cref="ArgumentException"/> rather than
    /// silently writing a literal cell value.</param>
    public Formula(string text)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (text.Length == 0)
            throw new ArgumentException("formula text must be non-empty", nameof(text));
        if (text[0] != '=')
            throw new ArgumentException(
                $"formula must start with '=': {text}", nameof(text));
        Text = text;
    }

    public bool Equals(Formula? other)
        => other is not null && string.Equals(Text, other.Text, StringComparison.Ordinal);

    public override bool Equals(object? obj) => Equals(obj as Formula);

    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Text);

    public override string ToString() => Text;
}
