using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace PyExcel.State;

/// <summary>
/// Pure-XML round-trip for the per-sheet <see cref="WorkbookProfileData"/> — the
/// workbook-scoped flags plus the <see cref="SheetProfile"/> map. Self-contained
/// (it does not depend on <see cref="WorkbookStateCodec"/>, which keeps owning the
/// older flat single-state format) so the per-sheet format can evolve on its own.
///
/// <para>The document is human-readable and versioned. The root carries
/// <c>xmlns="urn:pyexcel:workbook:1"</c>; <see cref="ProjectProfileCodec"/> nests
/// this element inside the workbook's profile part. Bumping the layout in a
/// non-backwards-compatible way means bumping the namespace.</para>
/// </summary>
public static class WorkbookProfileCodec
{
    /// <summary>XML namespace on the profile root.</summary>
    public const string XmlNamespace = "urn:pyexcel:workbook:1";

    /// <summary>Schema version attribute. Bump on a breaking layout change.</summary>
    public const string SchemaVersion = "1";

    private static readonly XNamespace Ns = XmlNamespace;

    // -------------------------------------------------------------------------
    // Serialize
    // -------------------------------------------------------------------------

    /// <summary>The <c>&lt;workbook&gt;</c> root element for
    /// <paramref name="data"/>, for nesting inside the project profile. Only
    /// configured sheets are emitted, in a stable (ordinal) key order so a no-op
    /// save doesn't churn the workbook's binary diff.</summary>
    public static XElement SerializeElement(WorkbookProfileData data)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));

        var root = new XElement(Ns + "workbook",
            new XAttribute("version", SchemaVersion),
            new XElement(Ns + "enabled", XmlConvert.ToString(data.Enabled)));

        if (data.ProjectDir is not null)
            root.Add(new XElement(Ns + "project-dir", data.ProjectDir));

        var sheets = new XElement(Ns + "sheets");
        foreach (var kv in data.Sheets.Where(kv => kv.Value.IsConfigured)
                                      .OrderBy(kv => kv.Key, StringComparer.Ordinal))
        {
            sheets.Add(SerializeSheet(kv.Key, kv.Value));
        }
        root.Add(sheets);

        return root;
    }

    /// <summary>Serialize <paramref name="data"/> to a standalone, indented XML
    /// string. Used by tests and for hand-inspection.</summary>
    public static string SerializeToString(WorkbookProfileData data)
        => new XDocument(new XDeclaration("1.0", "utf-8", null), SerializeElement(data))
            .ToString(SaveOptions.None);

    private static XElement SerializeSheet(string name, SheetProfile p)
    {
        var el = new XElement(Ns + "sheet", new XAttribute("name", name));
        AddText(el, "selected-script", p.SelectedScript);
        AddText(el, "py-input", p.PyInput);
        AddText(el, "py-output", p.PyOutput);
        AddText(el, "selected-action", p.SelectedActionName);
        AddText(el, "import-input", p.ImportInput);
        AddText(el, "import-output", p.ImportOutput);
        AddText(el, "export-input", p.ExportInput);
        AddText(el, "export-output", p.ExportOutput);
        AddText(el, "paste-output", p.PasteOutput);

        var actions = new XElement(Ns + "actions");
        foreach (var a in p.Actions) actions.Add(SerializeAction(a));
        el.Add(actions);
        return el;
    }

    private static XElement SerializeAction(RibbonAction a)
    {
        var el = new XElement(Ns + "action",
            new XAttribute("name", a.Name),
            new XAttribute("script", a.Script),
            new XAttribute("input", a.Input),
            new XAttribute("output", a.Output));

        if (a.Kwargs is { Count: > 0 })
        {
            var kwargs = new XElement(Ns + "kwargs");
            // Stable ordering keeps the on-disk format deterministic.
            foreach (var kv in a.Kwargs.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                kwargs.Add(new XElement(Ns + "kwarg",
                    new XAttribute("key", kv.Key),
                    new XAttribute("value", kv.Value)));
            }
            el.Add(kwargs);
        }

        return el;
    }

    // -------------------------------------------------------------------------
    // Deserialize
    // -------------------------------------------------------------------------

    /// <summary>Parse a <c>&lt;workbook&gt;</c> element produced by
    /// <see cref="SerializeElement"/>. Returns <see langword="false"/> (nulls out)
    /// for a null/foreign/unsupported element so a corrupt or foreign profile can
    /// never throw into a COM/ribbon caller.</summary>
    public static bool TryParseElement(XElement? root, out WorkbookProfileData? data)
    {
        data = null;
        if (root is null || root.Name != Ns + "workbook") return false;
        var version = (string?)root.Attribute("version");
        if (version != SchemaVersion) return false;

        try
        {
            var enabled = ParseBool((string?)root.Element(Ns + "enabled"));
            var projectDir = (string?)root.Element(Ns + "project-dir");

            var sheets = new Dictionary<string, SheetProfile>(StringComparer.Ordinal);
            var sheetsEl = root.Element(Ns + "sheets");
            if (sheetsEl is not null)
            {
                foreach (var s in sheetsEl.Elements(Ns + "sheet"))
                {
                    var name = (string?)s.Attribute("name") ?? string.Empty;
                    sheets[name] = ParseSheet(s);
                }
            }

            data = new WorkbookProfileData
            {
                Enabled = enabled,
                ProjectDir = projectDir,
                Sheets = sheets,
            };
            return true;
        }
        catch (Exception ex) when (ex is FormatException or XmlException)
        {
            data = null;
            return false;
        }
    }

    /// <summary>Best-effort parse of a standalone profile XML string.</summary>
    public static bool TryDeserialize(string? xml, out WorkbookProfileData? data)
    {
        data = null;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            return TryParseElement(XDocument.Parse(xml).Root, out data);
        }
        catch (Exception ex) when (ex is FormatException or XmlException)
        {
            return false;
        }
    }

    private static SheetProfile ParseSheet(XElement s) => new()
    {
        SelectedScript = (string?)s.Element(Ns + "selected-script"),
        PyInput = (string?)s.Element(Ns + "py-input"),
        PyOutput = (string?)s.Element(Ns + "py-output"),
        SelectedActionName = (string?)s.Element(Ns + "selected-action"),
        ImportInput = (string?)s.Element(Ns + "import-input"),
        ImportOutput = (string?)s.Element(Ns + "import-output"),
        ExportInput = (string?)s.Element(Ns + "export-input"),
        ExportOutput = (string?)s.Element(Ns + "export-output"),
        PasteOutput = (string?)s.Element(Ns + "paste-output"),
        Actions = ParseActions(s.Element(Ns + "actions")),
    };

    private static IReadOnlyList<RibbonAction> ParseActions(XElement? actionsEl)
    {
        if (actionsEl is null) return Array.Empty<RibbonAction>();
        var list = new List<RibbonAction>();
        foreach (var a in actionsEl.Elements(Ns + "action"))
            list.Add(ParseAction(a));
        return list;
    }

    private static RibbonAction ParseAction(XElement el)
    {
        var name = RequireAttribute(el, "name");
        var script = RequireAttribute(el, "script");
        var input = RequireAttribute(el, "input");
        var output = RequireAttribute(el, "output");

        Dictionary<string, string>? kwargs = null;
        var kwargsEl = el.Element(Ns + "kwargs");
        if (kwargsEl is not null)
        {
            kwargs = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var kv in kwargsEl.Elements(Ns + "kwarg"))
                kwargs[RequireAttribute(kv, "key")] = RequireAttribute(kv, "value");
        }

        return new RibbonAction(name, script, input, output, kwargs);
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static void AddText(XElement parent, string name, string? value)
    {
        if (value is not null) parent.Add(new XElement(Ns + name, value));
    }

    private static string RequireAttribute(XElement el, string name)
    {
        var attr = el.Attribute(name);
        if (attr is null)
            throw new FormatException(
                $"element '{el.Name.LocalName}' is missing required attribute '{name}'");
        return attr.Value;
    }

    private static bool ParseBool(string? value)
    {
        if (value is null) throw new FormatException("missing <enabled> element");
        try { return XmlConvert.ToBoolean(value.Trim().ToLowerInvariant()); }
        catch (FormatException ex) { throw new FormatException($"invalid <enabled> value '{value}'", ex); }
    }
}
