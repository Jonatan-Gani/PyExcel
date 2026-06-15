using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace PyExcel.State;

/// <summary>
/// Reads and writes a workbook's <see cref="ProjectProfile"/> as
/// <c>pyexcel.project.xml</c> in its project folder — the single, portable
/// source of truth for "is this workbook a PyExcel project, and what's in it?".
///
/// <para>This replaces the per-user app-data store: the project folder already
/// holds everything else (the venv, the kernel, userScripts, and — in PyExcel's
/// model — the workbook itself), so the state belongs there too. Copy or send
/// the folder and the project is intact, with no hidden per-machine data to also
/// copy. The portable in-file <c>CustomXMLPart</c> copy remains a fallback for
/// cloud workbooks that have no local folder.</para>
///
/// <para>On save, the environment metadata (OS, machine, CLR, add-in version,
/// and the project's Python path/version read from the venv's
/// <c>pyvenv.cfg</c>) is captured fresh, while <c>created-utc</c> and any
/// metadata that can't be recomputed are preserved from the existing file. All
/// operations are best-effort — I/O failures are swallowed so persistence can
/// never break a ribbon action or an Excel event.</para>
/// </summary>
public static class ProjectProfileStore
{
    /// <summary>Subfolder under the project dir holding PyExcel's internal
    /// files. Keeping the profile here (rather than a loose <c>.xml</c> next to
    /// the workbook) stops Excel from opening it as a sibling workbook, and
    /// groups it with the other dot-prefixed internals.</summary>
    public const string SubDirName = ".pyexcel";

    /// <summary>The profile filename, written inside <see cref="SubDirName"/>.</summary>
    public const string FileName = "project.xml";

    // The venv directory Setup provisions (mirrors PyExcel.Setup.VenvProvisioner
    // / PyExcel.Excel.PythonResolver). Duplicated as a literal here to avoid a
    // project reference just to read a folder name.
    private const string VenvDirName = ".pyexcel-venv";

    /// <summary>Full path to the profile file for <paramref name="projectDir"/>,
    /// or null if no project dir was given.</summary>
    public static string? PathFor(string? projectDir)
        => string.IsNullOrWhiteSpace(projectDir)
            ? null
            : Path.Combine(projectDir!, SubDirName, FileName);

    /// <summary>The loose profile path an earlier build wrote (a bare
    /// <c>pyexcel.project.xml</c> next to the workbook), kept only for read
    /// fallback + cleanup.</summary>
    private static string? LegacyPathFor(string? projectDir)
        => string.IsNullOrWhiteSpace(projectDir)
            ? null
            : Path.Combine(projectDir!, "pyexcel.project.xml");

    private static void TryDeleteLegacy(string projectDir)
    {
        try
        {
            var legacy = LegacyPathFor(projectDir);
            if (legacy is not null && File.Exists(legacy)) File.Delete(legacy);
        }
        catch
        {
            // Best-effort cleanup.
        }
    }

    /// <summary>Write the profile for <paramref name="state"/> into
    /// <paramref name="projectDir"/>, capturing fresh environment metadata and
    /// preserving non-recomputable fields from any existing file.</summary>
    public static void Save(string? projectDir, WorkbookState state, string? workbookName, string? workbookPath)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));
        if (string.IsNullOrWhiteSpace(projectDir)) return;
        try
        {
            var file = PathFor(projectDir)!;
            Directory.CreateDirectory(Path.GetDirectoryName(file)!);
            var prior = TryLoadProfile(projectDir, state.WorkbookKey)?.Metadata;
            var meta = BuildMetadata(projectDir!, workbookName, workbookPath, prior);
            File.WriteAllText(file, ProjectProfileCodec.SerializeToString(state, meta), new UTF8Encoding(false));

            // Remove the legacy loose profile next to the workbook (earlier builds
            // wrote it there); Excel would otherwise keep offering to open it as a
            // workbook. Best-effort.
            TryDeleteLegacy(projectDir!);
        }
        catch
        {
            // Best-effort: never let persistence break the caller.
        }
    }

    /// <summary>Load just the workbook state from the project profile, or null
    /// if there's no readable profile in <paramref name="projectDir"/>.</summary>
    public static WorkbookState? TryLoad(string? projectDir, string workbookKey)
        => TryLoadProfile(projectDir, workbookKey)?.State;

    /// <summary>Load the full profile (state + metadata) from
    /// <paramref name="projectDir"/>, or null if none is readable. Falls back to
    /// the legacy loose location an earlier build used, so already-enabled
    /// projects keep working (and get migrated to the subfolder on next save).</summary>
    public static ProjectProfile? TryLoadProfile(string? projectDir, string workbookKey)
    {
        return ReadFrom(PathFor(projectDir), workbookKey)
               ?? ReadFrom(LegacyPathFor(projectDir), workbookKey);
    }

    private static ProjectProfile? ReadFrom(string? path, string workbookKey)
    {
        if (path is null) return null;
        try
        {
            if (!File.Exists(path)) return null;
            if (ProjectProfileCodec.TryDeserialize(File.ReadAllText(path), workbookKey, out var state, out var meta)
                && state is not null)
            {
                return new ProjectProfile(state, meta ?? new ProjectMetadata());
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static ProjectMetadata BuildMetadata(
        string projectDir, string? workbookName, string? workbookPath, ProjectMetadata? prior)
    {
        var (pythonPath, pythonVersion) = ReadVenvPython(projectDir);
        var now = DateTimeOffset.UtcNow;
        return new ProjectMetadata(
            GeneratedBy: "PyExcel " + AssemblyVersion(),
            CreatedUtc: prior?.CreatedUtc ?? now,
            UpdatedUtc: now,
            Os: Safe(() => RuntimeInformation.OSDescription),
            Machine: Safe(() => Environment.MachineName),
            ProcessBits: Environment.Is64BitProcess ? 64 : 32,
            Clr: Safe(() => RuntimeInformation.FrameworkDescription),
            // Fresh from the venv if we can read it; otherwise keep what we had.
            PythonPath: pythonPath ?? prior?.PythonPath,
            PythonVersion: pythonVersion ?? prior?.PythonVersion,
            WorkbookName: workbookName ?? prior?.WorkbookName,
            WorkbookPath: workbookPath ?? prior?.WorkbookPath,
            ProjectDir: projectDir);
    }

    /// <summary>Read the project's Python executable path and version from the
    /// venv's <c>pyvenv.cfg</c> (cheap, no process spawn). Returns nulls if the
    /// venv isn't there yet or can't be read.</summary>
    private static (string? path, string? version) ReadVenvPython(string projectDir)
    {
        try
        {
            var venv = Path.Combine(projectDir, VenvDirName);
            var cfg = Path.Combine(venv, "pyvenv.cfg");
            string? version = null;
            if (File.Exists(cfg))
            {
                foreach (var raw in File.ReadAllLines(cfg))
                {
                    var line = raw.Trim();
                    var eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    var keyName = line.Substring(0, eq).Trim();
                    if (keyName.Equals("version", StringComparison.OrdinalIgnoreCase)
                        || keyName.Equals("version_info", StringComparison.OrdinalIgnoreCase))
                    {
                        version = line.Substring(eq + 1).Trim();
                        break;
                    }
                }
            }

            // Windows venvs put the interpreter in Scripts\; POSIX in bin/.
            var win = Path.Combine(venv, "Scripts", "python.exe");
            var posix = Path.Combine(venv, "bin", "python");
            string? path = File.Exists(win) ? win : File.Exists(posix) ? posix : null;

            return (path, version);
        }
        catch
        {
            return (null, null);
        }
    }

    private static string AssemblyVersion()
        => typeof(ProjectProfileStore).Assembly.GetName().Version?.ToString() ?? "unknown";

    private static string? Safe(Func<string?> f)
    {
        try { return f(); }
        catch { return null; }
    }
}
