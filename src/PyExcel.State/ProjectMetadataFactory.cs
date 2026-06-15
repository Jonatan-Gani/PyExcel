using System;
using System.IO;
using System.Runtime.InteropServices;

namespace PyExcel.State;

/// <summary>
/// Builds the <see cref="ProjectMetadata"/> provenance block embedded alongside
/// the user <see cref="WorkbookState"/> in a workbook's PyExcel profile. Pure and
/// cross-platform — it only reads files and process/runtime info, no COM — so it
/// builds on the netstandard slice and is unit-tested on Linux.
///
/// <para>On each save the environment fields (OS, machine, CLR, add-in version,
/// and the project's Python path/version read cheaply from the venv's
/// <c>pyvenv.cfg</c>) are captured fresh, while <c>created-utc</c> and any field
/// that can't be recomputed are preserved from the prior metadata already on the
/// workbook.</para>
/// </summary>
public static class ProjectMetadataFactory
{
    // The venv directory Setup provisions (mirrors PyExcel.Setup.VenvProvisioner
    // / PyExcel.Excel.PythonResolver). Duplicated as a literal here to avoid a
    // project reference just to read a folder name.
    private const string VenvDirName = ".pyexcel-venv";

    /// <summary>Capture fresh environment metadata for <paramref name="projectDir"/>,
    /// preserving <see cref="ProjectMetadata.CreatedUtc"/> and any non-recomputable
    /// field from <paramref name="prior"/> (the metadata already embedded in the
    /// workbook, if any).</summary>
    public static ProjectMetadata Build(
        string? projectDir, string? workbookName, string? workbookPath, ProjectMetadata? prior)
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
            ProjectDir: projectDir ?? prior?.ProjectDir);
    }

    /// <summary>Read the project's Python executable path and version from the
    /// venv's <c>pyvenv.cfg</c> (cheap, no process spawn). Returns nulls if no
    /// project dir is known, or the venv isn't there yet / can't be read.</summary>
    private static (string? path, string? version) ReadVenvPython(string? projectDir)
    {
        if (string.IsNullOrWhiteSpace(projectDir)) return (null, null);
        try
        {
            var venv = Path.Combine(projectDir!, VenvDirName);
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
        => typeof(ProjectMetadataFactory).Assembly.GetName().Version?.ToString() ?? "unknown";

    private static string? Safe(Func<string?> f)
    {
        try { return f(); }
        catch { return null; }
    }
}
