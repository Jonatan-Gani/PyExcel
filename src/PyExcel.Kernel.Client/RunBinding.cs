namespace PyExcel.Kernel.Client;

/// <summary>
/// One input or output binding declared on a run request: the name the
/// value appears under in the kernel's <c>inputs</c> dict, and the Python
/// type the kernel must construct (inputs) or assert (outputs).
///
/// <para><see cref="Type"/> is the lowercase wire token from
/// <c>PyExcel.State.PyExcelTypes.WireName</c> — carried as a plain string
/// so this assembly stays free of a dependency on the state layer, exactly
/// as it stays free of Apache.Arrow and just shuttles bytes. The
/// authoritative vocabulary lives in <c>docs/typed-io-contract.md</c> and
/// is mirrored by <c>embedded/pyexcel/kernel/declared_types.py</c>.</para>
///
/// <para>An input binding's type must already be resolved — the host
/// measures the range, so <c>auto</c> is settled before the request is
/// built. An output binding may legitimately carry <c>auto</c>, which
/// means "do not enforce".</para>
/// </summary>
public sealed class RunBinding
{
    /// <summary>The binding's name. <see langword="null"/> or empty leaves
    /// the kernel to auto-name it by resolved type and ordinal
    /// (<c>df1</c>, <c>list1</c>, <c>value1</c>, …).</summary>
    public string? Name { get; }

    /// <summary>The declared type's wire token (e.g. <c>"dataframe"</c>).</summary>
    public string Type { get; }

    /// <summary>The binding's range text, used only to make the kernel's
    /// error messages name the range the user actually configured.</summary>
    public string? Range { get; }

    public RunBinding(string? name, string type, string? range = null)
    {
        Name = name;
        Type = string.IsNullOrWhiteSpace(type) ? "auto" : type;
        Range = range;
    }
}
