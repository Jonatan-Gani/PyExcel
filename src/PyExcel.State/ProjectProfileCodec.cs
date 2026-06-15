using System;
using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace PyExcel.State;

/// <summary>
/// XML round-trip for the <see cref="ProjectProfile"/> — the document embedded in
/// the workbook's <c>CustomXMLPart</c> by <c>WorkbookStatePersister</c>. It is
/// deliberately human-readable (indented, one element per field) so a user or
/// support engineer who extracts the part can see exactly what the project is:
/// its metadata block plus the per-sheet workbook profile.
///
/// <para>The user-profile portion nests <see cref="WorkbookProfileCodec"/>'s
/// <c>&lt;workbook&gt;</c> element (its own <c>urn:pyexcel:workbook:1</c>
/// namespace) inside the profile root. For backwards compatibility it also reads
/// an earlier build's nested flat single-state document
/// (<c>{urn:pyexcel:state:1}pyexcel</c>) and migrates it forward via
/// <see cref="WorkbookProfileData.FromState"/>, so an already-saved workbook keeps
/// its configuration.</para>
/// </summary>
public static class ProjectProfileCodec
{
    /// <summary>Namespace on the profile root, distinct from the nested profile.</summary>
    public const string XmlNamespace = "urn:pyexcel:project:1";

    /// <summary>Profile schema version. Bump on a breaking layout change.</summary>
    public const string SchemaVersion = "1";

    private static readonly XNamespace Ns = XmlNamespace;

    /// <summary>The earlier-build flat single-state root we still read (and migrate
    /// forward) so already-saved workbooks don't lose their configuration.</summary>
    private static readonly XName FlatStateRootName =
        XNamespace.Get(WorkbookStateCodec.XmlNamespace) + "pyexcel";

    /// <summary>Serialize <paramref name="data"/> + <paramref name="meta"/> to an
    /// indented, human-readable XML string.</summary>
    public static string SerializeToString(WorkbookProfileData data, ProjectMetadata meta)
        => Serialize(data, meta).ToString(SaveOptions.None);

    public static XDocument Serialize(WorkbookProfileData data, ProjectMetadata meta)
    {
        if (data is null) throw new ArgumentNullException(nameof(data));
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

        var root = new XElement(Ns + "pyexcel-project",
            new XAttribute("project-version", SchemaVersion),
            metadata,
            WorkbookProfileCodec.SerializeElement(data));

        return new XDocument(new XDeclaration("1.0", "utf-8", null), root);
    }

    /// <summary>Best-effort parse. Returns <see langword="false"/> (nulls out) for
    /// null/blank/non-XML input or anything that isn't a recognised profile — so a
    /// corrupt or foreign file never throws into a COM/ribbon caller. The
    /// caller-supplied <paramref name="workbookKey"/> is used only when migrating
    /// an earlier-build flat document (whose nested codec keyed by it).</summary>
    public static bool TryDeserialize(
        string? xml, string workbookKey, out WorkbookProfileData? data, out ProjectMetadata? meta)
    {
        if (workbookKey is null) throw new ArgumentNullException(nameof(workbookKey));
        data = null;
        meta = null;
        if (string.IsNullOrWhiteSpace(xml)) return false;
        try
        {
            var doc = XDocument.Parse(xml);
            var root = doc.Root;
            if (root is null || root.Name != Ns + "pyexcel-project") return false;

            // Current format: the nested per-sheet <workbook> element.
            if (WorkbookProfileCodec.TryParseElement(
                    root.Element(XNamespace.Get(WorkbookProfileCodec.XmlNamespace) + "workbook"),
                    out var parsed)
                && parsed is not null)
            {
                data = parsed;
                meta = ReadMetadata(root.Element(Ns + "metadata"));
                return true;
            }

            // Earlier build: a nested flat single-state document — migrate forward.
            var flat = root.Element(FlatStateRootName);
            if (flat is not null)
            {
                var state = WorkbookStateCodec.Deserialize(new XDocument(new XElement(flat)), workbookKey);
                data = WorkbookProfileData.FromState(state);
                meta = ReadMetadata(root.Element(Ns + "metadata"));
                return true;
            }

            return false;
        }
        catch (Exception ex) when (ex is FormatException or XmlException)
        {
            data = null;
            meta = null;
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
