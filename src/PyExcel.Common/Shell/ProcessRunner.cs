using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using PyExcel.Common.Logging;

// The argv quoter (QuoteArgs) is an internal implementation detail of
// the CommandLineToArgvW round-trip, but its correctness is worth
// testing directly — the edge cases (trailing backslashes, embedded
// quotes) are exactly the kind that break silently. Expose internals to
// the test assembly, same pattern PyExcel.Bridge uses for its tests.
[assembly: InternalsVisibleTo("PyExcel.Bridge.Tests")]

namespace PyExcel.Common.Shell;

/// <summary>
/// Cross-platform child-process runner that captures stdout and stderr
/// to an <see cref="ILog"/> while accumulating the full streams for the
/// caller. Used by <c>PyExcel.Setup</c> for venv creation, pip install,
/// and dependency probes — every external command Setup issues funnels
/// through here so a single log file (<c>%TEMP%\PyExcel_Debug.log</c>
/// by default) tells the operator exactly what ran and what came back.
///
/// <para>Contract:</para>
/// <list type="bullet">
///   <item>Never invokes a shell. <see cref="ProcessStartInfo.UseShellExecute"/>
///     is fixed to <c>false</c>; arguments are an explicit <c>string[]</c>
///     that gets quoted with the Windows CommandLineToArgvW rules.</item>
///   <item>Both streams are read concurrently to avoid the classic
///     pipe-buffer deadlock (stderr fills, child blocks on write, parent
///     blocks on stdout read).</item>
///   <item>Each line is forwarded to <see cref="ILog"/> with stdout →
///     <see cref="ILog.Info"/> and stderr → <see cref="ILog.Warn"/>; the
///     full text is also retained on the returned <see cref="ProcessRunResult"/>
///     for callers that want to inspect or surface it.</item>
///   <item>On timeout, the child is killed (best effort) and a
///     <see cref="TimeoutException"/> is thrown after the already-captured
///     output is flushed to the log.</item>
/// </list>
///
/// <para><b>Why not <c>PyExcel.Common.Shell.ShellLauncher</c>:</b> that
/// helper opens files with their registered shell handler
/// (<c>UseShellExecute=true</c>) — opposite semantics to what an
/// automation runner needs. Both wrappers belong here, but they
/// deliberately don't share an implementation.</para>
///
/// <para><b>TFM note:</b> <c>ProcessStartInfo.ArgumentList</c> is
/// netstandard2.1+; we target netstandard2.0 so we build the command
/// line ourselves. The quoter below implements the documented
/// CommandLineToArgvW round-trip rules so a path with spaces, embedded
/// quotes, or trailing backslashes survives parsing on the child side.
/// On POSIX, <see cref="Process"/> on .NET uses the same
/// <see cref="ProcessStartInfo.Arguments"/> string but most callers pass
/// plain paths without metacharacters, so the same quoting rules apply
/// in practice; quoting a UNC or spaces-containing path with double
/// quotes is safe for python, pip, and venv invocations.</para>
/// </summary>
public sealed class ProcessRunner
{
    private readonly ILog _log;

    public ProcessRunner(ILog? log = null)
    {
        _log = log ?? NullLog.Instance;
    }

    /// <summary>
    /// Run <paramref name="fileName"/> with the given argv list. Blocks
    /// until the child exits or <paramref name="timeoutMs"/> elapses.
    /// </summary>
    /// <param name="fileName">Absolute path to the executable, or a bare
    ///     name to be resolved via PATH by the OS.</param>
    /// <param name="args">Argument list. Each entry is passed as a single
    ///     argv element to the child — no shell tokenisation.</param>
    /// <param name="workingDirectory">Optional working directory; the
    ///     child inherits the parent's CWD when null.</param>
    /// <param name="environment">Optional environment overlay; entries
    ///     are merged on top of the parent's environment. Pass
    ///     <see langword="null"/> to inherit the parent environment as-is.</param>
    /// <param name="timeoutMs">Wall-clock timeout. A 10-minute default
    ///     matches a typical pip-install upper bound on a slow network;
    ///     callers with tighter SLAs should pass their own.</param>
    /// <param name="cancellationToken">Cancellation token. When tripped,
    ///     the child is killed and a <see cref="OperationCanceledException"/>
    ///     surfaces to the caller after captured output is flushed.</param>
    /// <returns>Exit code plus the full stdout and stderr the child wrote.</returns>
    /// <exception cref="ArgumentException">file name is null/whitespace.</exception>
    /// <exception cref="TimeoutException">child did not exit within
    ///     <paramref name="timeoutMs"/>; the partial output captured so
    ///     far has been logged.</exception>
    /// <exception cref="OperationCanceledException">the cancellation
    ///     token tripped while the child was still running.</exception>
    public ProcessRunResult Run(
        string fileName,
        IReadOnlyList<string> args,
        string? workingDirectory = null,
        IReadOnlyDictionary<string, string>? environment = null,
        int timeoutMs = 600_000,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("file name required", nameof(fileName));
        if (args is null) throw new ArgumentNullException(nameof(args));
        if (timeoutMs <= 0) throw new ArgumentOutOfRangeException(nameof(timeoutMs));

        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = QuoteArgs(args),
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = false,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        if (!string.IsNullOrWhiteSpace(workingDirectory))
            psi.WorkingDirectory = workingDirectory!;
        if (environment is { })
        {
            foreach (var kv in environment)
                psi.Environment[kv.Key] = kv.Value;
        }

        var argDisplay = ArgvForLog(fileName, args);
        _log.Info($"exec: {argDisplay}");
        if (!string.IsNullOrWhiteSpace(workingDirectory))
            _log.Debug($"  cwd: {workingDirectory!}");

        var stdout = new StringBuilder();
        var stderr = new StringBuilder();

        using var proc = new Process { StartInfo = psi };
        proc.OutputDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            lock (stdout) stdout.AppendLine(e.Data);
            _log.Info($"  {e.Data}");
        };
        proc.ErrorDataReceived += (_, e) =>
        {
            if (e.Data is null) return;
            lock (stderr) stderr.AppendLine(e.Data);
            _log.Warn($"  {e.Data}");
        };

        if (!proc.Start())
            throw new InvalidOperationException(
                $"failed to start process: {fileName}");

        proc.BeginOutputReadLine();
        proc.BeginErrorReadLine();

        var waited = WaitWithCancellation(proc, timeoutMs, cancellationToken);
        if (waited == WaitOutcome.Cancelled)
        {
            TryKill(proc);
            // After Kill returns, drain the async stdout/stderr readers
            // with the parameterless WaitForExit so any buffered output
            // the child managed to flush before dying still lands in the
            // log. The bounded-timeout overload does NOT wait for the
            // BeginOutput/ErrorReadLine handlers — only the no-arg form
            // does — so we follow Kill with a final no-arg wait.
            try { proc.WaitForExit(); } catch { /* ignore */ }
            _log.Warn($"process cancelled after partial run: {argDisplay}");
            cancellationToken.ThrowIfCancellationRequested();
        }
        if (waited == WaitOutcome.TimedOut)
        {
            TryKill(proc);
            try { proc.WaitForExit(); } catch { /* ignore */ }
            _log.Error($"process timed out after {timeoutMs}ms: {argDisplay}");
            throw new TimeoutException(
                $"process {fileName} did not exit within {timeoutMs}ms");
        }

        // WaitForExit(int) returns true the moment the process exits but
        // does NOT block until the async readers have flushed their
        // final lines. The parameterless WaitForExit does — call it
        // unconditionally here so stdout/stderr are complete before we
        // construct the result.
        try { proc.WaitForExit(); } catch { /* ignore */ }

        var exitCode = proc.ExitCode;
        _log.Info($"exit {exitCode}: {argDisplay}");
        return new ProcessRunResult(exitCode, stdout.ToString(), stderr.ToString());
    }

    // -------------------------------------------------------------------------
    // Internals
    // -------------------------------------------------------------------------

    private enum WaitOutcome { Exited, TimedOut, Cancelled }

    private static WaitOutcome WaitWithCancellation(
        Process proc, int timeoutMs, CancellationToken cancellationToken)
    {
        if (!cancellationToken.CanBeCanceled)
            return proc.WaitForExit(timeoutMs) ? WaitOutcome.Exited : WaitOutcome.TimedOut;

        // Poll in short slices so the cancellation token has a bounded
        // response time. The cost is negligible — venv/pip runs are
        // measured in seconds, the slice is 100 ms.
        const int sliceMs = 100;
        var remaining = timeoutMs;
        while (remaining > 0)
        {
            if (cancellationToken.IsCancellationRequested)
                return WaitOutcome.Cancelled;
            var slice = Math.Min(sliceMs, remaining);
            if (proc.WaitForExit(slice))
                return WaitOutcome.Exited;
            remaining -= slice;
        }
        return WaitOutcome.TimedOut;
    }

    private static void TryKill(Process p)
    {
        try { if (!p.HasExited) p.Kill(); }
        catch { /* already dead or insufficient perms */ }
    }

    /// <summary>
    /// Quote an argv list into the single string ProcessStartInfo.Arguments
    /// wants, following the documented Windows CommandLineToArgvW rules:
    /// backslashes are literal except when they precede a double-quote;
    /// double-quotes inside an argument are escaped with a backslash; the
    /// whole argument is wrapped in double-quotes if it contains
    /// whitespace, a quote, or is empty.
    /// </summary>
    internal static string QuoteArgs(IReadOnlyList<string> args)
    {
        var sb = new StringBuilder();
        for (var i = 0; i < args.Count; i++)
        {
            if (i > 0) sb.Append(' ');
            QuoteSingle(args[i], sb);
        }
        return sb.ToString();
    }

    private static void QuoteSingle(string arg, StringBuilder sb)
    {
        if (arg is null)
            throw new ArgumentException("argv entry must not be null");

        // Empty argument: must be preserved as an empty quoted string,
        // otherwise CommandLineToArgvW drops it entirely.
        if (arg.Length == 0)
        {
            sb.Append("\"\"");
            return;
        }

        // Fast path: argument has no whitespace, no quote, no backslash;
        // pass through verbatim.
        var needsQuoting = false;
        foreach (var c in arg)
        {
            if (c == ' ' || c == '\t' || c == '\n' || c == '\v' || c == '"')
            {
                needsQuoting = true;
                break;
            }
        }
        if (!needsQuoting)
        {
            sb.Append(arg);
            return;
        }

        sb.Append('"');
        var backslashRun = 0;
        foreach (var c in arg)
        {
            if (c == '\\')
            {
                backslashRun++;
                continue;
            }
            if (c == '"')
            {
                // Each pending backslash AND the quote itself must be
                // backslash-escaped because the quote terminates the
                // backslash run inside a quoted argument.
                sb.Append('\\', backslashRun * 2 + 1);
                backslashRun = 0;
                sb.Append('"');
                continue;
            }
            // Normal character — the pending backslashes are literal.
            sb.Append('\\', backslashRun);
            backslashRun = 0;
            sb.Append(c);
        }
        // Trailing backslashes at the end of the argument must be doubled
        // because the closing quote that follows would otherwise eat them.
        sb.Append('\\', backslashRun * 2);
        sb.Append('"');
    }

    private static string ArgvForLog(string fileName, IReadOnlyList<string> args)
    {
        // The same quoting we pass to the child, prefixed by the file name,
        // so the log line is paste-runnable for an operator debugging a
        // failed install.
        var sb = new StringBuilder();
        QuoteSingle(fileName, sb);
        for (var i = 0; i < args.Count; i++)
        {
            sb.Append(' ');
            QuoteSingle(args[i], sb);
        }
        return sb.ToString();
    }
}

/// <summary>
/// Outcome of a <see cref="ProcessRunner.Run(string, IReadOnlyList{string}, string?, IReadOnlyDictionary{string, string}?, int, CancellationToken)"/>
/// call. <see cref="ExitCode"/> = 0 conventionally means success; non-zero
/// is failure with <see cref="Stderr"/> usually carrying the diagnostic.
/// </summary>
public sealed class ProcessRunResult
{
    public int ExitCode { get; }
    public string Stdout { get; }
    public string Stderr { get; }

    public ProcessRunResult(int exitCode, string stdout, string stderr)
    {
        ExitCode = exitCode;
        Stdout = stdout ?? string.Empty;
        Stderr = stderr ?? string.Empty;
    }

    /// <summary>Convenience predicate: did the child exit with status 0.</summary>
    public bool Success => ExitCode == 0;
}
