using System;
using System.Collections.Generic;

namespace PyExcel.Excel;

/// <summary>
/// One named result from a <c>transform</c> that returned a dict.
/// <see cref="Value"/> is already decoded into its spill-ready CLR shape —
/// <c>object?[,]</c> for a table or vector, a boxed scalar, a
/// <see cref="ChartSpec"/>, or a <see cref="ChartImage"/>.
/// </summary>
public sealed class PyRunOutput
{
    /// <summary>The dict key this result came back under, or
    /// <see langword="null"/> when the kernel reported fewer names than
    /// payloads.</summary>
    public string? Name { get; }

    /// <summary>The decoded, spill-ready value.</summary>
    public object Value { get; }

    public PyRunOutput(string? name, object value)
    {
        Name = name;
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }
}

/// <summary>
/// The decoded result of a run whose <c>transform</c> returned a dict of
/// named results.
///
/// <para>This is what makes the documented multi-output contract real. The
/// frame format always length-prefixed a payload count, but the host read
/// only the first payload, so a dict return either lost every key past the
/// first or — before the kernel learned to split it — arrived as a single
/// Arrow struct that reached the sheet as the literal string
/// <c>"StructArray"</c>.</para>
///
/// <para>Callers that can route by name (the ribbon's Run button, via the
/// Output field's bindings) write each entry to its own range. Callers that
/// cannot — the single-cell <c>=PY.RUN</c> UDF — take the sole entry when
/// there is exactly one and report a clear error otherwise.</para>
/// </summary>
public sealed class PyRunOutputs
{
    /// <summary>The named results, in the order the user's dict enumerated.</summary>
    public IReadOnlyList<PyRunOutput> Outputs { get; }

    public PyRunOutputs(IReadOnlyList<PyRunOutput> outputs)
        => Outputs = outputs ?? throw new ArgumentNullException(nameof(outputs));

    /// <summary>Look up one result by key, case-sensitively (Python dict
    /// keys are case-sensitive, so matching them loosely would invent
    /// behaviour the script does not have).</summary>
    public PyRunOutput? ByName(string name)
    {
        foreach (var output in Outputs)
            if (string.Equals(output.Name, name, StringComparison.Ordinal))
                return output;
        return null;
    }
}
