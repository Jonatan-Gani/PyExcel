using System;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace PyExcel.State;

/// <summary>
/// XML round-trip for the <see cref="ProjectProfile"/> — the project-folder
/// file (<c>pyexcel.project.xml</c>). The document is deliberately human-
/// readable (indented, one element per field) so a user or support engineer can
/// open it and see exactly what the project is: its metadata block plus the
/// user state.
///
/// <para>The user-state portion reuses the already-tested
/// <see cref="WorkbookStateCodec"/> verbatim — the profile simply nests that
/// document's <c>&lt;pyexcel&gt;</c> element (in its own <c>urn:pyexcel:state:1</c>
/// namespace) inside the profile root, so there's one serializer for state and
/// no risk of the two drifting.</para>
/// </summary>
public static class ProjectProfileCodec
{
    /// <summary>Namespace on the profile root, distinct from the state
    /// namespace it nests.</summary>
    public const string XmlNamespace = "urn:pyexcel:project:1";

    /// <summary>Profile schema version. Bump on a breaking layout change.</summary>
    public const string SchemaVersion = "1";

    private static readonly XNamespace Ns = XmlNamespace;
    private static readonly XName StateRootName =
        XNamespace.Get(WorkbookStateCodec.XmlNamespace) + "pyexcel";

    /// <summary>Serialize <paramref name="state"/> + <paramref name="meta"/> to
    /// an indented, human-readable XML string.</summary>
    public static string SerializeToString(WorkbookState state, ProjectMetadata meta)
        => Serialize(state, meta).ToString(SaveOptions.None);

    public static XDocument Serialize(WorkbookState state, ProjectMetadata meta)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));
        if (meta is null) throw new ArgumentNullException(nameof(meta));

        var metadata = new XElement(Ns + "metadata");
        Add(metadata, "generated-by", meta.GeneratedBy);
        Add(metadata, "created-utc", meta.CreatedUtc?.ToString("o", CultureInfo.InvariantCulture));
        Add(metadata, "updated-utc", meta.UpdatedUtc?.ToString("o", CultureInfo.InvariantCulture));
        Add(metadata, "os", meta.Os);
        Add(metadata, "machine", meta.Machine);
        Add(metadata, "process-bits", meta.ProcessBits?.ToString(CultureInfo.InvariantCulture));
        Add(metadata, "clr", meta.Clr);
        Add(metadata, "python-path", meta.PythonPath);
        Add(metadata, "python-version", meta.PythonVersion);
        Add(metadata, "workbook-name", meta.WorkbookName);
        Add(metadata, "workbook-path", meta.WorkbookPath);
        Add(metadata, "project-dir", meta.ProjectDir);

        // Reuse the tested state serializer and nest a copy of its root element
        // (copy so it's detached from its own document before re-parenting).
        var stateRoot = new XElement(WorkbookStateCodec.Serialize(state).Root!);

        var root = new XElement(Ns + "pyexcel-project",
            new XAttribute("project-version", SchemaVersion),
            metadata,
            stateRoot);

        return new XDocument(new XDeclaration("1.0", "utf-8", null), root);
    }

    /// <summary>Best-effort parse. Returns <see langword="false"/> (nulls out)
    /// for null/blank/non-XML input or anything that isn't a recognised profile
    /// — so a corrupt or foreign file never throws into a COM/ribbon caller. The
    /// caller-supplied <paramref name="key"/> always wins over the file.</summary>
    public static bool TryDeserialize(
        string? xml, string workbookKey, out WorkbookState? state, out ProjectMetadata? meta)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        state = null;
        meta = null;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            var doc = XDocument.Parse(xml);
            var root = doc.Root;
            if (root is null || root.Name != Ns + "pyexcel-project") return false;

            var stateRoot = root.Element(StateRootName);
            if (stateRoot is null) return false;

            // Clone the nested state element into its own document for the state codec.
            state = WorkbookStateCodec.Deserialize(new XDocument(new XElement(stateRoot)), workbookKey);
            meta = ReadMetadata(root.Element(Ns + "metadata"));
            return true;
        }
        catch (Exception ex) when (ex is FormatException or XmlException)
        {
            return false;
        }
    }

    private static ProjectMetadata ReadMetadata(XElement? m)
    {
        if (m is null) return new ProjectMetadata();
        return new ProjectMetadata(
            GeneratedBy: Str(m, "generated-by"),
            CreatedUtc: Date(m, "created-utc"),
            UpdatedUtc: Date(m, "updated-utc"),
            Os: Str(m, "os"),
            Machine: Str(m, "machine"),
            ProcessBits: Int(m, "process-bits"),
            Clr: Str(m, "clr"),
            PythonPath: Str(m, "python-path"),
            PythonVersion: Str(m, "python-version"),
            WorkbookName: Str(m, "workbook-name"),
            WorkbookPath: Str(m, "workbook-path"),
            ProjectDir: Str(m, "project-dir"));
    }

    private static void Add(XElement parent, string name, string? value)
    {
        if (!string.IsNullOrEmpty(value)) parent.Add(new XElement(Ns + name, value));
    }

    private static string? Str(XElement m, string name) => (string?)m.Element(Ns + name);

    private static int? Int(XElement m, string name)
        => int.TryParse(Str(m, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    private static DateTimeOffset? Date(XElement m, string name)
        => DateTimeOffset.TryParse(
            Str(m, name), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var v) ? v : null;
}
