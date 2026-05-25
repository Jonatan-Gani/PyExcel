using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace PyExcel.State;

/// <summary>
/// Watches a directory for <c>.py</c> files appearing, disappearing, or
/// being renamed, and pushes the updated script list to a caller-supplied
/// callback. Used by Phase 3 to keep the ribbon's Script dropdown live
/// against whatever the user has stashed in <c>userScripts/</c>.
///
/// <para>The callback fires synchronously on the
/// <see cref="FileSystemWatcher"/> worker thread. If the caller has UI
/// constraints (e.g. <c>IRibbonUI.Invalidate</c> must run on the main
/// Excel thread), it is the caller's job to queue the work back onto
/// the right thread — typically via Excel-DNA's
/// <c>ExcelAsyncUtil.QueueAsMacro</c>.</para>
///
/// <para>The constructor pushes an initial snapshot before returning so
/// callers don't have to special-case the "watcher just started, no
/// changes yet" state.</para>
/// </summary>
public sealed class ScriptDirectoryWatcher : IDisposable
{
    private readonly FileSystemWatcher _watcher;
    private readonly Action<IReadOnlyList<string>> _onScriptsChanged;
    private readonly string _directory;
    private int _disposed;

    public ScriptDirectoryWatcher(string directory, Action<IReadOnlyList<string>> onScriptsChanged)
    {
        if (directory is null) throw new ArgumentNullException(nameof(directory));
        if (onScriptsChanged is null) throw new ArgumentNullException(nameof(onScriptsChanged));
        if (!Directory.Exists(directory))
            throw new DirectoryNotFoundException(directory);

        _directory = directory;
        _onScriptsChanged = onScriptsChanged;

        _watcher = new FileSystemWatcher(directory, "*.py")
        {
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
            IncludeSubdirectories = false,
            EnableRaisingEvents = false,  // turned on after handlers attach
        };
        _watcher.Created += OnFsEvent;
        _watcher.Deleted += OnFsEvent;
        _watcher.Renamed += OnFsRenamed;
        _watcher.EnableRaisingEvents = true;

        // Initial snapshot so consumers don't sit at "no scripts" until
        // the first edit lands.
        PushSnapshot();
    }

    /// <summary>Stop watching and release the FileSystemWatcher.</summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        try { _watcher.EnableRaisingEvents = false; } catch { /* swallow */ }
        try { _watcher.Dispose(); } catch { /* swallow */ }
    }

    /// <summary>Re-scan the directory and push a fresh list. Public so
    /// callers can force a refresh after a coarse event (e.g. workbook
    /// activation) the watcher itself doesn't see.</summary>
    public void Refresh() => PushSnapshot();

    private void OnFsEvent(object sender, FileSystemEventArgs e) => PushSnapshot();
    private void OnFsRenamed(object sender, RenamedEventArgs e) => PushSnapshot();

    private void PushSnapshot()
    {
        if (Volatile.Read(ref _disposed) != 0) return;

        IReadOnlyList<string> snapshot;
        try
        {
            snapshot = Directory.EnumerateFiles(_directory, "*.py")
                .Select(Path.GetFileNameWithoutExtension)
                .Where(n => !string.IsNullOrEmpty(n))
                .Select(n => n!)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (DirectoryNotFoundException)
        {
            // Directory deleted out from under us — surface as empty list
            // so the ribbon clears rather than throwing.
            snapshot = Array.Empty<string>();
        }
        catch (IOException)
        {
            // Transient FS errors (e.g. an antivirus scanning a file we
            // tried to enumerate). Don't crash the watcher thread; the
            // next event will retry the snapshot.
            return;
        }

        try { _onScriptsChanged(snapshot); }
        catch { /* user callback threw; not our concern */ }
    }
}
