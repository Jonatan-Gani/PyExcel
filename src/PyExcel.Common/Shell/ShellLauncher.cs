#if NETFRAMEWORK
using System;
using System.Diagnostics;
using System.IO;

namespace PyExcel.Common.Shell;

/// <summary>
/// Thin wrapper over the Windows shell verbs the ribbon needs to launch
/// external programs: open a file with its registered handler, reveal a
/// file in Explorer, open a folder. Used by the Phase 5 ribbon
/// callbacks (<c>OnReadMe</c>, <c>OnOpenExplorer</c>, <c>OnEditPython</c>)
/// to delegate to whatever the user has set as their default editor /
/// file manager rather than encoding tool choice into PyExcel.
///
/// <para>Net48-only on purpose — <see cref="ProcessStartInfo.UseShellExecute"/>
/// has materially different semantics across .NET versions and OS
/// targets; PyExcel only ever runs as a 64-bit Excel-DNA <c>.xll</c> on
/// Windows, so we don't pretend to be portable here.</para>
///
/// <para>Each method validates its input (throws
/// <see cref="ArgumentException"/> for null/empty paths) but does not
/// attempt to verify that the path exists — that's the caller's
/// responsibility, because the right fallback (silently no-op vs. show a
/// "missing file" dialog) is policy that lives with the ribbon
/// callback, not in this helper.</para>
/// </summary>
public static class ShellLauncher
{
    /// <summary>Open <paramref name="path"/> with its registered shell
    /// handler — Markdown viewer for <c>.md</c>, default Python editor
    /// for <c>.py</c>, file manager for a directory, and so on. The user
    /// configures these via the OS's Default Apps settings; PyExcel
    /// doesn't second-guess them.</summary>
    /// <exception cref="ArgumentException"><paramref name="path"/> is
    /// null, empty, or whitespace.</exception>
    public static void Open(string path)
    {
        RequirePath(path, nameof(path));
        Process.Start(new ProcessStartInfo
        {
            FileName = path,
            UseShellExecute = true,
        });
    }

    /// <summary>Reveal <paramref name="path"/> in Windows Explorer. If
    /// the target is a file, Explorer opens the parent directory with
    /// the file selected; if it's a directory, Explorer opens that
    /// directory directly. Uses the documented <c>/select,</c> verb for
    /// the file case (note the comma — required by Explorer's
    /// argument-parser quirk).</summary>
    /// <exception cref="ArgumentException"><paramref name="path"/> is
    /// null, empty, or whitespace.</exception>
    public static void OpenInExplorer(string path)
    {
        RequirePath(path, nameof(path));

        if (Directory.Exists(path))
        {
            // Plain directory open — explorer.exe takes a single arg.
            Process.Start(new ProcessStartInfo
            {
                FileName = "explorer.exe",
                Arguments = QuoteArg(path),
                UseShellExecute = true,
            });
            return;
        }

        // Either a file or a non-existent path. Use /select, so an
        // existing file is highlighted; for a missing path Explorer
        // falls back to opening the parent directory, which is the
        // sensible behaviour for "the file isn't there yet."
        Process.Start(new ProcessStartInfo
        {
            FileName = "explorer.exe",
            Arguments = $"/select,{QuoteArg(path)}",
            UseShellExecute = true,
        });
    }

    /// <summary>Quote an argument for the legacy <c>explorer.exe</c>
    /// command line. Explorer doesn't understand the standard
    /// CommandLineToArgvW escape rules — passing a path with spaces
    /// without quotes is what the documentation says to do, but it
    /// actually fails on UNC paths. Wrapping in double quotes is the
    /// reliable form across local paths, UNC paths, and paths with
    /// spaces; embedded double quotes in a path are not legal on
    /// Windows so we don't escape them.</summary>
    private static string QuoteArg(string arg)
        => "\"" + arg + "\"";

    private static void RequirePath(string path, string name)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("path is null or empty", name);
    }
}
#endif
