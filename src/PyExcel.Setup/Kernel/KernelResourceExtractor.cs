using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using PyExcel.Common.Logging;

namespace PyExcel.Setup.Kernel;

/// <summary>
/// Extracts the embedded <c>pyexcel</c> Python package and the canonical
/// <c>requirements.txt</c> from this assembly's manifest resources onto
/// disk, where the kernel subprocess can <c>import pyexcel.kernel</c>
/// from a PYTHONPATH-discoverable directory.
///
/// <para><b>Why this replaces v1's Setup.bas pipeline:</b> the v1
/// add-in stored the kernel sources base64-encoded in a hidden sheet
/// and reassembled them by concatenating 32 000-character chunks at
/// install time. That coupled deployment to Excel's storage limits,
/// made source diffs unreadable in version control, and routinely
/// corrupted on cross-version round-trips. The v2 replacement is a
/// plain <c>&lt;EmbeddedResource&gt;</c> per file (see
/// <c>PyExcel.Setup.csproj</c>) read out at runtime with the standard
/// <see cref="Assembly.GetManifestResourceStream(string)"/> API — no
/// base64, no chunk assembly, no sheet.</para>
///
/// <para>Logical-resource layout (set by <c>LogicalName</c> in the
/// csproj):</para>
/// <list type="bullet">
///   <item><c>pyexcel/__init__.py</c></item>
///   <item><c>pyexcel/kernel/__init__.py</c></item>
///   <item><c>pyexcel/kernel/__main__.py</c></item>
///   <item><c>pyexcel/kernel/*.py</c> (one entry per kernel module)</item>
///   <item><c>pyexcel/requirements.txt</c></item>
/// </list>
///
/// <para>Extraction is idempotent: an existing file is only overwritten
/// when its bytes differ from the embedded copy, and the
/// <c>last-write</c> timestamp is preserved when no change is needed.
/// That lets a user keep a pinned kernel checkout in place across
/// add-in updates if they're deliberately running a forked kernel,
/// while still picking up shipped updates the moment the bytes change.</para>
/// </summary>
public sealed class KernelResourceExtractor
{
    /// <summary>
    /// Manifest-resource prefix every kernel file shares. Used to
    /// distinguish the kernel payload from any future unrelated
    /// resources baked into this assembly.
    /// </summary>
    public const string ResourcePrefix = "pyexcel/";

    /// <summary>
    /// Name of the canonical requirements file inside the resource set.
    /// Surfaced as a constant so <see cref="Pip.PipRunner"/> and the
    /// dependency verifier can pull the same string.
    /// </summary>
    public const string RequirementsLogicalName = "pyexcel/requirements.txt";

    private readonly Assembly _source;
    private readonly ILog _log;

    /// <summary>Construct against an explicit source assembly (tests
    /// substitute a fixture assembly here) and logger.</summary>
    public KernelResourceExtractor(Assembly? source = null, ILog? log = null)
    {
        _source = source ?? typeof(KernelResourceExtractor).Assembly;
        _log = log ?? NullLog.Instance;
    }

    /// <summary>
    /// Extract every <c>pyexcel/…</c> resource onto <paramref name="targetDir"/>,
    /// preserving the full logical-name path so the resulting tree
    /// matches what <c>import pyexcel.kernel</c> expects on PYTHONPATH:
    /// <c>targetDir/pyexcel/__init__.py</c>,
    /// <c>targetDir/pyexcel/kernel/__main__.py</c>, etc. The caller
    /// adds <paramref name="targetDir"/> (not the
    /// <c>targetDir/pyexcel</c> subdirectory) to PYTHONPATH.
    /// Sub-directories are created as needed; existing files are
    /// overwritten only when content differs.
    /// </summary>
    /// <returns>The list of relative paths written or refreshed (empty
    /// when nothing changed).</returns>
    /// <exception cref="ArgumentException">target dir is null/whitespace.</exception>
    /// <exception cref="InvalidOperationException">no kernel resources
    /// are embedded in the source assembly — typically a build
    /// configuration regression that dropped the <c>&lt;EmbeddedResource&gt;</c>
    /// items.</exception>
    public ExtractionResult Extract(string targetDir)
    {
        if (string.IsNullOrWhiteSpace(targetDir))
            throw new ArgumentException("target directory required", nameof(targetDir));

        var resources = EnumerateKernelResources().ToList();
        if (resources.Count == 0)
            throw new InvalidOperationException(
                $"no kernel resources found on {_source.FullName}; " +
                $"check PyExcel.Setup.csproj <EmbeddedResource> items.");

        Directory.CreateDirectory(targetDir);

        var written = new List<string>();
        var skipped = new List<string>();
        foreach (var name in resources)
        {
            // Preserve the full logical-name path: a resource named
            // `pyexcel/kernel/__main__.py` lands at
            // `targetDir/pyexcel/kernel/__main__.py`. We do NOT strip
            // the `pyexcel/` prefix — Python needs the package
            // directory to exist on disk for `import pyexcel.kernel`
            // to work.
            var relative = name.Replace('/', Path.DirectorySeparatorChar);
            var path = Path.Combine(targetDir, relative);
            var parent = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(parent))
                Directory.CreateDirectory(parent);

            var bytes = ReadResource(name);

            if (File.Exists(path) && BytesEqual(File.ReadAllBytes(path), bytes))
            {
                skipped.Add(relative);
                continue;
            }

            File.WriteAllBytes(path, bytes);
            written.Add(relative);
            _log.Debug($"extracted {relative} ({bytes.Length} bytes)");
        }

        _log.Info(
            $"kernel extraction complete: {written.Count} written, " +
            $"{skipped.Count} up-to-date, target={targetDir}");

        return new ExtractionResult(targetDir, written, skipped);
    }

    /// <summary>
    /// Report whether the kernel already on disk at
    /// <paramref name="targetDir"/> matches the one embedded in this build,
    /// without writing anything.
    ///
    /// <para>This is the read-only half of <see cref="Extract(string)"/> and
    /// uses the same byte comparison, so the two can never disagree about what
    /// "up to date" means.</para>
    ///
    /// <para><b>Content, not version.</b> Comparing <c>__version__</c> would be
    /// cheaper but is unsound: the kernel's version string only changes when
    /// someone remembers to change it, so a build that altered kernel code
    /// without bumping it reads as identical. That exact case shipped — a
    /// rewritten worker.py under an unchanged <c>2.0.0a0</c> — and left users
    /// running an old kernel against a new host with no signal at all. Bytes
    /// cannot be forgotten.</para>
    /// </summary>
    /// <returns>A check naming every resource that is missing or differs.</returns>
    /// <exception cref="ArgumentException">target dir is null/whitespace.</exception>
    public KernelExtractionCheck Check(string targetDir)
    {
        if (string.IsNullOrWhiteSpace(targetDir))
            throw new ArgumentException("target directory required", nameof(targetDir));

        var stale = new List<string>();

        foreach (var name in EnumerateKernelResources())
        {
            var relative = name.Replace('/', Path.DirectorySeparatorChar);
            var path = Path.Combine(targetDir, relative);

            if (!File.Exists(path))
            {
                stale.Add(relative);
                continue;
            }

            try
            {
                if (!BytesEqual(File.ReadAllBytes(path), ReadResource(name)))
                    stale.Add(relative);
            }
            catch (IOException)
            {
                // Unreadable is not the same as absent, but for the caller's
                // purposes — "should this be re-extracted?" — it is.
                stale.Add(relative);
            }
        }

        return new KernelExtractionCheck(targetDir, stale);
    }

    /// <summary>
    /// Enumerate the logical names of every kernel resource embedded in
    /// the source assembly, in deterministic order. Public so tests can
    /// assert the expected shipping set without running an extraction.
    /// </summary>
    public IEnumerable<string> EnumerateKernelResources()
    {
        return _source
            .GetManifestResourceNames()
            .Where(n => n.StartsWith(ResourcePrefix, StringComparison.Ordinal))
            .OrderBy(n => n, StringComparer.Ordinal);
    }

    private byte[] ReadResource(string logicalName)
    {
        using var stream = _source.GetManifestResourceStream(logicalName);
        if (stream is null)
            throw new InvalidOperationException(
                $"resource '{logicalName}' missing from {_source.FullName}");
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    private static bool BytesEqual(byte[] a, byte[] b)
    {
        if (a.Length != b.Length) return false;
        for (var i = 0; i < a.Length; i++)
            if (a[i] != b[i]) return false;
        return true;
    }
}

/// <summary>
/// Result of a <see cref="KernelResourceExtractor.Check(string)"/> — whether
/// the kernel on disk still matches the one this build ships.
///
/// <para><see cref="Stale"/> lists the relative paths that are missing or
/// differ; it is empty exactly when <see cref="UpToDate"/> is true.</para>
/// </summary>
public sealed class KernelExtractionCheck
{
    /// <summary>The directory that was inspected.</summary>
    public string TargetDir { get; }

    /// <summary>Relative paths that are missing from disk or differ from the
    /// embedded copy.</summary>
    public IReadOnlyList<string> Stale { get; }

    /// <summary>True when every embedded kernel file is present on disk with
    /// identical bytes.</summary>
    public bool UpToDate => Stale.Count == 0;

    public KernelExtractionCheck(string targetDir, IReadOnlyList<string> stale)
    {
        TargetDir = targetDir ?? throw new ArgumentNullException(nameof(targetDir));
        Stale = stale ?? throw new ArgumentNullException(nameof(stale));
    }

    /// <summary>A one-line summary for the log, naming a few offenders so a
    /// reader can tell a wholly-missing kernel from a single changed file.</summary>
    public string Describe()
    {
        if (UpToDate) return $"kernel at '{TargetDir}' matches this build";
        var sample = string.Join(", ", Stale.Count > 3 ? Stale.Take(3) : Stale);
        var more = Stale.Count > 3 ? $" (+{Stale.Count - 3} more)" : "";
        return $"kernel at '{TargetDir}' is out of date: {Stale.Count} file(s) differ — {sample}{more}";
    }
}

/// <summary>
/// Summary of a <see cref="KernelResourceExtractor.Extract(string)"/>
/// call. <see cref="Written"/> and <see cref="Skipped"/> are relative
/// paths under <see cref="TargetDir"/>.
/// </summary>
public sealed class ExtractionResult
{
    public string TargetDir { get; }
    public IReadOnlyList<string> Written { get; }
    public IReadOnlyList<string> Skipped { get; }

    public ExtractionResult(string targetDir, IReadOnlyList<string> written, IReadOnlyList<string> skipped)
    {
        TargetDir = targetDir;
        Written = written;
        Skipped = skipped;
    }
}
