using System;

namespace PyExcel.Excel;

/// <summary>
/// A JSON chart specification produced by the kernel when a user
/// <c>transform()</c> returns a Plotly figure. The host parses it with
/// <see cref="ChartSpecParser"/> and renders a native Excel chart via
/// <c>ChartBuilder</c>.
///
/// <para>The wire representation is a 1×1 Arrow string column whose
/// schema carries <c>pyexcel-shape = chart</c>. See
/// <see cref="ArrowMarshal"/> for the decode plumbing and
/// <c>embedded/pyexcel/kernel/chart.py</c> for the spec schema and the
/// matching Python dataclass.</para>
/// </summary>
public sealed class ChartSpec : IEquatable<ChartSpec>
{
    /// <summary>The serialised chart-spec document.</summary>
    public string Json { get; }

    public ChartSpec(string json)
    {
        if (json is null) throw new ArgumentNullException(nameof(json));
        if (string.IsNullOrWhiteSpace(json))
            throw new ArgumentException("chart spec JSON must be non-empty", nameof(json));
        Json = json;
    }

    public bool Equals(ChartSpec? other)
        => other is not null && string.Equals(Json, other.Json, StringComparison.Ordinal);

    public override bool Equals(object? obj) => Equals(obj as ChartSpec);

    public override int GetHashCode() => StringComparer.Ordinal.GetHashCode(Json);

    public override string ToString() => "[PyExcel chart spec]";
}
