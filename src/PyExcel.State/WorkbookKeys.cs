using System;

namespace PyExcel.State;

/// <summary>
/// Shared workbook-key derivation. Both the <c>ExcelWorkbookContext</c>
/// (which answers "what workbook is active right now") and the COM event
/// sink (which fires on open / activate / save / close for a <em>specific</em>
/// workbook) must compute the same key for the same workbook — otherwise
/// state set through one path is invisible to the other, closed workbooks
/// leak, and save-on-close writes the wrong part.
///
/// <para>The rule mirrors the original <c>ExcelWorkbookContext</c>
/// strategy:</para>
/// <list type="bullet">
///   <item>Saved workbook → its <c>FullName</c> (the full on-disk path).
///     Stable across close/reopen because the path is the workbook's
///     identity.</item>
///   <item>Unsaved workbook (empty <c>Path</c>) → a synthetic
///     <c>"unsaved:{SessionGuid}:{Name}"</c> key. Excel hands every new
///     workbook a session-unique <c>Name</c> (<c>Book1</c>, <c>Book2</c>,
///     …); the session GUID — allocated once per add-in load — keeps this
///     session's unsaved books from colliding with anything else.</item>
/// </list>
/// </summary>
public static class WorkbookKeys
{
    /// <summary>One GUID per add-in load, shared by every component that
    /// derives an unsaved-workbook key so the keys agree process-wide.</summary>
    public static readonly string SessionGuid = Guid.NewGuid().ToString("N");

    /// <summary>Key for an unsaved workbook identified only by its
    /// (session-unique) <paramref name="name"/>.</summary>
    public static string UnsavedKey(string name)
    {
        if (name is null) throw new ArgumentNullException(nameof(name));
        return $"unsaved:{SessionGuid}:{name}";
    }

    /// <summary>Resolve a workbook key from the three COM properties the
    /// caller reads off the workbook. A workbook with no on-disk
    /// <paramref name="path"/> is treated as unsaved and keyed by
    /// <paramref name="name"/>; otherwise the <paramref name="fullName"/>
    /// (the full path) is the key.</summary>
    public static string Resolve(string name, string path, string fullName)
    {
        if (name is null) throw new ArgumentNullException(nameof(name));
        if (fullName is null) throw new ArgumentNullException(nameof(fullName));
        return string.IsNullOrEmpty(path) ? UnsavedKey(name) : fullName;
    }
}
