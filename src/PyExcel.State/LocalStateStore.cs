using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace PyExcel.State;

/// <summary>
/// Reliable per-user persistence of a workbook's PyExcel state, keyed by the
/// workbook's on-disk path, stored under the local app-data folder
/// (<c>%LOCALAPPDATA%\PyExcel\workbooks</c>).
///
/// <para>This is the durable answer to "is <em>this</em> workbook a PyExcel
/// project, and what's in it?" — written the instant the user enables a
/// workbook or edits an action, and read back when the workbook is opened or
/// re-activated. It deliberately does <b>not</b> depend on Excel's save /
/// <c>CustomXMLPart</c> round-trip, which only persists if the user saves and
/// the <c>BeforeSave</c> event fires and the part survives the write — too many
/// failure points for the "I enabled it, why is it asking again?" experience.
/// The <c>CustomXMLPart</c> (via <c>WorkbookStatePersister</c>) stays as the
/// portable copy that travels with the file to another machine; this store is
/// the reliable local source of truth.</para>
///
/// <para>Only workbooks with a real path are persisted — an unsaved workbook's
/// key (<c>unsaved:{guid}:{name}</c>) isn't stable across sessions, so there's
/// nothing durable to key on until it's saved. All operations are best-effort:
/// I/O failures are swallowed so persistence can never break a ribbon action.</para>
/// </summary>
public static class LocalStateStore
{
    /// <summary>Folder holding one XML file per known workbook.</summary>
    public static string Root
    {
        get
        {
            var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var baseDir = string.IsNullOrEmpty(localAppData)
                ? Path.Combine(Path.GetTempPath(), "PyExcel")
                : Path.Combine(localAppData, "PyExcel");
            return Path.Combine(baseDir, "workbooks");
        }
    }

    /// <summary>Persist <paramref name="state"/> for <paramref name="key"/>.
    /// No-op for unsaved-workbook keys. Best-effort.</summary>
    public static void Save(string key, WorkbookState state)
    {
        if (state is null) throw new ArgumentNullException(nameof(state));
        if (!IsPersistableKey(key)) return;
        try
        {
            Directory.CreateDirectory(Root);
            File.WriteAllText(PathFor(key), WorkbookStateCodec.SerializeToString(state), new UTF8Encoding(false));
        }
        catch
        {
            // Best-effort: never let persistence break the action that triggered it.
        }
    }

    /// <summary>Load the stored state for <paramref name="key"/>, or
    /// <see langword="null"/> if none is stored / it's unreadable / the key
    /// isn't persistable.</summary>
    public static WorkbookState? TryLoad(string key)
    {
        if (!IsPersistableKey(key)) return null;
        try
        {
            var path = PathFor(key);
            if (!File.Exists(path)) return null;
            return WorkbookStateCodec.TryDeserialize(File.ReadAllText(path), key, out var state)
                ? state
                : null;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>Drop the stored state for <paramref name="key"/>. Best-effort;
    /// used when a workbook is un-enabled (not yet wired, but the symmetric
    /// operation belongs here).</summary>
    public static void Remove(string key)
    {
        if (!IsPersistableKey(key)) return;
        try
        {
            var path = PathFor(key);
            if (File.Exists(path)) File.Delete(path);
        }
        catch
        {
            // Best-effort.
        }
    }

    /// <summary>A key is persistable iff it's a real workbook path — i.e. not
    /// the synthetic <c>unsaved:{guid}:{name}</c> key an unsaved workbook gets.</summary>
    private static bool IsPersistableKey(string? key)
        => !string.IsNullOrEmpty(key) && !key!.StartsWith("unsaved:", StringComparison.Ordinal);

    private static string PathFor(string key) => Path.Combine(Root, Hash(key) + ".xml");

    /// <summary>Stable, filesystem-safe filename for a workbook path. The full
    /// path itself can't be a filename (it contains separators), so hash it.
    /// SHA-256 hex — no <c>Convert.ToHexString</c> on netstandard2.0, so format
    /// via <see cref="BitConverter"/>.</summary>
    private static string Hash(string key)
    {
        using var sha = SHA256.Create();
        var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(key));
        return BitConverter.ToString(bytes).Replace("-", string.Empty);
    }
}
