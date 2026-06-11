using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using PyExcel.Common.Logging;
using PyExcel.Common.Shell;
using Xunit;

namespace PyExcel.Bridge.Tests;

/// <summary>
/// Behavioural tests for <see cref="ProcessRunner"/> plus the
/// CommandLineToArgvW round-trip quoter that backs it. Cross-platform:
/// the quoter tests run on every lane; the process-execution tests
/// pick a host-appropriate command (echo / cmd.exe).
/// </summary>
public class ProcessRunnerTests
{
    // -------------------------------------------------------------------------
    // QuoteArgs — pure CommandLineToArgvW round-trip rules
    // -------------------------------------------------------------------------

    [Fact]
    public void QuoteArgs_EmptyList_ReturnsEmptyString()
    {
        Assert.Equal(string.Empty, ProcessRunner.QuoteArgs(Array.Empty<string>()));
    }

    [Fact]
    public void QuoteArgs_PlainArgs_NoQuoting()
    {
        var quoted = ProcessRunner.QuoteArgs(new[] { "-m", "pip", "install" });
        Assert.Equal("-m pip install", quoted);
    }

    [Fact]
    public void QuoteArgs_ArgWithSpace_IsWrappedInQuotes()
    {
        var quoted = ProcessRunner.QuoteArgs(new[] { @"C:\Program Files\Python\python.exe" });
        Assert.Equal(@"""C:\Program Files\Python\python.exe""", quoted);
    }

    [Fact]
    public void QuoteArgs_EmptyArg_BecomesEmptyQuotedString()
    {
        var quoted = ProcessRunner.QuoteArgs(new[] { "a", "", "b" });
        Assert.Equal(@"a """" b", quoted);
    }

    [Fact]
    public void QuoteArgs_ArgWithEmbeddedQuote_EscapesQuote()
    {
        // Per CommandLineToArgvW rules the embedded `"` becomes `\"`,
        // and the whole arg is wrapped in `"`.
        var quoted = ProcessRunner.QuoteArgs(new[] { @"say ""hi""" });
        Assert.Equal(@"""say \""hi\""""", quoted);
    }

    [Fact]
    public void QuoteArgs_ArgEndingInBackslash_DoublesTrailingBackslashes()
    {
        // `C:\foo bar\` (note the space requiring quoting) ends in one
        // backslash; with the closing quote following, that one
        // backslash must be doubled so the parser does not interpret
        // it as escaping the close-quote.
        var quoted = ProcessRunner.QuoteArgs(new[] { @"C:\foo bar\" });
        Assert.Equal(@"""C:\foo bar\\""", quoted);
    }

    [Fact]
    public void QuoteArgs_BackslashesBeforeQuote_AreDoubled()
    {
        // Two backslashes before a literal quote: each backslash is
        // doubled and the quote itself is backslash-escaped, giving
        // five characters (4 backslashes + escaped quote).
        var quoted = ProcessRunner.QuoteArgs(new[] { @"a\\""b" });
        Assert.Equal(@"""a\\\\\""b""", quoted);
    }

    [Fact]
    public void QuoteArgs_InternalBackslashes_StayLiteral()
    {
        // No quoting needed at all — backslashes that don't precede a
        // quote or a closing-quote position are literal.
        var quoted = ProcessRunner.QuoteArgs(new[] { @"C:\Users\you\file.txt" });
        Assert.Equal(@"C:\Users\you\file.txt", quoted);
    }

    // -------------------------------------------------------------------------
    // Run — host-specific commands prove the streams and exit codes work
    // -------------------------------------------------------------------------

    [Fact]
    public void Run_ExitZero_ReportsSuccessAndCapturesStdout()
    {
        var (file, args) = HelloWorld();
        var result = new ProcessRunner().Run(file, args);
        Assert.True(result.Success, $"expected success, stderr={result.Stderr}");
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("hello", result.Stdout, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Run_NonZeroExit_IsReportedNotThrown()
    {
        var (file, args) = ExitsWithCode(7);
        var result = new ProcessRunner().Run(file, args);
        Assert.False(result.Success);
        Assert.Equal(7, result.ExitCode);
    }

    [Fact]
    public void Run_Timeout_KillsChildAndThrows()
    {
        var (file, args) = LongRunning();
        var ex = Assert.Throws<TimeoutException>(() =>
            new ProcessRunner().Run(file, args, timeoutMs: 250));
        Assert.Contains("did not exit", ex.Message);
    }

    [Fact]
    public void Run_Cancellation_KillsChildAndThrows()
    {
        var (file, args) = LongRunning();
        using var cts = new CancellationTokenSource();
        cts.CancelAfter(150);
        Assert.Throws<OperationCanceledException>(() =>
            new ProcessRunner().Run(file, args, timeoutMs: 30_000, cancellationToken: cts.Token));
    }

    [Fact]
    public void Run_LogsLinesToILog()
    {
        var (file, args) = HelloWorld();
        var capture = new CapturingLog();
        var result = new ProcessRunner(capture).Run(file, args);
        Assert.True(result.Success);
        Assert.Contains(capture.InfoLines, line => line.Contains("hello", StringComparison.OrdinalIgnoreCase));
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private static (string file, string[] args) HelloWorld()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return ("cmd.exe", new[] { "/c", "echo hello" });
        return ("/bin/sh", new[] { "-c", "echo hello" });
    }

    private static (string file, string[] args) ExitsWithCode(int code)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return ("cmd.exe", new[] { "/c", $"exit {code}" });
        return ("/bin/sh", new[] { "-c", $"exit {code}" });
    }

    private static (string file, string[] args) LongRunning()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            // ping -n 60 = ~60 seconds of waiting; far longer than any test timeout.
            return ("cmd.exe", new[] { "/c", "ping -n 60 127.0.0.1 > NUL" });
        return ("/bin/sh", new[] { "-c", "sleep 60" });
    }

    private sealed class CapturingLog : ILog
    {
        public System.Collections.Generic.List<string> InfoLines { get; } = new();
        public System.Collections.Generic.List<string> WarnLines { get; } = new();
        public System.Collections.Generic.List<string> ErrorLines { get; } = new();

        public void Trace(string message) { }
        public void Debug(string message) { }
        public void Info(string message) => InfoLines.Add(message);
        public void Warn(string message) => WarnLines.Add(message);
        public void Error(string message, Exception? exception = null) => ErrorLines.Add(message);
    }
}
