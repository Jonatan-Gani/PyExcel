using System;
using System.IO;
using System.Text;

namespace PyExcel.Common.Logging;

/// <summary>
/// One-line-per-event file logger backed by an append-only text file.
///
/// Mirrors the v1 <c>LogToFile</c> behaviour in
/// <c>src/module/HostManager.bas:332-338</c> so that an operator who knows
/// where the v1 log lives can continue to use the same path. By default we
/// write to <c>%TEMP%\PyExcel_Debug.log</c>; the constructor accepts an
/// override for tests.
///
/// Writes use a short-lived <see cref="StreamWriter"/> per call. That is
/// intentionally pessimistic: in a .xll, we can be invoked from arbitrary
/// COM-thread contexts including during shutdown, and we would rather pay a
/// per-call open cost than risk a held handle leaking past <c>AutoClose</c>.
/// Failures are swallowed — the logger never throws into the caller, by
/// design (matching v1's <c>On Error Resume Next</c>).
/// </summary>
public sealed class FileLog : ILog
{
    private readonly string _path;
    private readonly object _gate = new();

    public FileLog(string? path = null)
    {
        _path = path ?? DefaultPath();
    }

    public static string DefaultPath()
    {
        var temp = Environment.GetEnvironmentVariable("TEMP")
                   ?? Path.GetTempPath();
        return Path.Combine(temp, "PyExcel_Debug.log");
    }

    public void Trace(string message) => Write("TRACE", message);
    public void Debug(string message) => Write("DEBUG", message);
    public void Info(string message)  => Write("INFO ", message);
    public void Warn(string message)  => Write("WARN ", message);

    public void Error(string message, Exception? exception = null)
    {
        if (exception is null)
        {
            Write("ERROR", message);
        }
        else
        {
            Write("ERROR", $"{message} :: {exception.GetType().Name}: {exception.Message}");
        }
    }

    private void Write(string level, string message)
    {
        // Build the line outside the lock so we hold it for the shortest
        // possible window. Single string concat — no StringBuilder needed
        // for a sub-200-byte line.
        var line = string.Concat(
            "[", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"), "] [",
            level, "] ", message);

        try
        {
            lock (_gate)
            {
                using var fs = new FileStream(
                    _path, FileMode.Append, FileAccess.Write, FileShare.Read);
                using var sw = new StreamWriter(fs, Encoding.UTF8);
                sw.WriteLine(line);
            }
        }
        catch
        {
            // Logging must never throw at the call site. Swallow per the
            // class-level contract.
        }
    }
}
