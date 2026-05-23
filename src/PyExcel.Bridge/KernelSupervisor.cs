using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

[assembly: InternalsVisibleTo("PyExcel.Kernel.Client")]
[assembly: InternalsVisibleTo("PyExcel.Bridge.Tests")]

namespace PyExcel.Bridge;

/// <summary>
/// Owns the lifecycle of one Python kernel subprocess.
///
/// Flow on construction:
///   1. Create a server-side <see cref="NamedPipeServerStream"/> with a
///      caller-supplied unique name.
///   2. Spawn <c>python -m pyexcel.kernel --pipe &lt;name&gt;</c> with an
///      explicit argv array (never a shell command string).
///   3. Wait for the child to connect, with timeout. Bail with a clear
///      error if the child exits before connecting.
///   4. HELLO handshake — server sends HELLO first, client replies with
///      HELLO carrying its protocol version. Mismatch fails the handshake.
///
/// <para>Concurrency model:</para>
/// <list type="bullet">
///   <item><c>ExchangeSemaphore</c> serialises high-level request/response
///     sequences — only one of <see cref="Ping"/>, <see cref="Shutdown"/>,
///     or a <c>PyExcel.Kernel.Client.KernelClient.Run</c> runs at a time.</item>
///   <item><c>WriteLock</c> guards each individual <c>WriteFrame</c> call
///     so byte writes never interleave on the pipe.</item>
///   <item><c>ReadLock</c> guards each individual <c>ReadFrame</c> call.</item>
/// </list>
/// <para>The split lets fire-and-forget frames (the CANCEL frame sent by
/// <c>KernelClient.Cancel</c>) acquire only the write lock — so they can
/// fire while a <c>Run</c> is parked in a read.</para>
///
/// Disposal is the canonical cleanup path: best-effort SHUTDOWN frame,
/// short wait for the child to exit, then <see cref="Process.Kill()"/> if
/// it didn't. Always destroys the pipe and the child process — no orphaned
/// python.exe even if the C# side throws mid-handshake.
/// </summary>
public sealed class KernelSupervisor : IDisposable
{
    private readonly Process _process;
    private readonly NamedPipeServerStream _pipe;
    private readonly FrameTransport _transport;
    private readonly SemaphoreSlim _exchange = new(1, 1);
    private readonly object _readLock = new();
    private readonly object _writeLock = new();
    private bool _disposed;

    // Internal accessors used by PyExcel.Kernel.Client. The split-lock
    // pattern means callers must lock the right lock for each call:
    // WriteLock for any WriteFrame, ReadLock for any ReadFrame, and
    // ExchangeSemaphore around any multi-frame request/response sequence.
    internal FrameTransport Transport => _transport;
    internal object ReadLock => _readLock;
    internal object WriteLock => _writeLock;
    internal SemaphoreSlim ExchangeSemaphore => _exchange;

    public string PipeName { get; }
    public Process Process => _process;
    public int RemoteProtocolVersion { get; }

    private KernelSupervisor(
        Process process,
        NamedPipeServerStream pipe,
        FrameTransport transport,
        string pipeName,
        int remoteProtocolVersion)
    {
        _process = process;
        _pipe = pipe;
        _transport = transport;
        PipeName = pipeName;
        RemoteProtocolVersion = remoteProtocolVersion;
    }

    /// <summary>
    /// Spawn the kernel and complete the HELLO handshake.
    ///
    /// <paramref name="pythonPath"/> is prepended to PYTHONPATH for the
    /// child so it can <c>import pyexcel.kernel</c> regardless of where
    /// the kernel package lives on disk (in production, alongside the
    /// .xll; in tests, the repo's <c>embedded/</c> directory).
    /// </summary>
    public static KernelSupervisor StartPython(
        string pythonExecutable,
        string pythonPath,
        int handshakeTimeoutMs = 10_000,
        int maxFrameBytes = Framing.DefaultMaxFrameBytes)
    {
        if (string.IsNullOrWhiteSpace(pythonExecutable))
            throw new ArgumentException("python executable required", nameof(pythonExecutable));
        if (string.IsNullOrWhiteSpace(pythonPath))
            throw new ArgumentException("PYTHONPATH directory required", nameof(pythonPath));
        if (handshakeTimeoutMs <= 0)
            throw new ArgumentOutOfRangeException(nameof(handshakeTimeoutMs));

        var pipeName = "pyexcel-kernel-" + Guid.NewGuid().ToString("N");

        // PipeOptions.Asynchronous: required for overlapped I/O on Windows
        // even when we use the synchronous Read/Write APIs; without it, a
        // concurrent ReadFrame + WriteFrame on the same pipe deadlocks.
        var pipe = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            transmissionMode: PipeTransmissionMode.Byte,
            options: PipeOptions.Asynchronous);

        Process? proc = null;
        try
        {
            var connectTask = pipe.WaitForConnectionAsync();
            proc = StartChild(pythonExecutable, pythonPath, pipeName);

            // Race the connection against the child exiting early. If the
            // child dies before connecting, surface its stderr — otherwise
            // callers see an opaque "timed out" with no clue what broke.
            WaitForConnectOrExit(connectTask, proc, handshakeTimeoutMs);

            var transport = new FrameTransport(pipe, ownsStream: false, maxFrameBytes: maxFrameBytes);
            var protocolVersion = Handshake(transport, handshakeTimeoutMs);

            return new KernelSupervisor(proc, pipe, transport, pipeName, protocolVersion);
        }
        catch
        {
            // Don't leak pipe handles or child processes if anything in the
            // handshake path throws. KillAndDispose is best-effort.
            TryKill(proc);
            try { pipe.Dispose(); } catch { /* swallow */ }
            throw;
        }
    }

    /// <summary>
    /// Send a PING and wait for the matching PONG. Returns the round-trip
    /// duration. Throws <see cref="TimeoutException"/> if no reply arrives
    /// within <paramref name="timeoutMs"/>.
    /// </summary>
    public TimeSpan Ping(int timeoutMs = 2000)
    {
        ThrowIfDisposed();
        if (timeoutMs <= 0) throw new ArgumentOutOfRangeException(nameof(timeoutMs));

        if (!_exchange.Wait(timeoutMs))
            throw new TimeoutException($"another exchange held the kernel for {timeoutMs}ms");
        try
        {
            var nonce = Guid.NewGuid().ToString("N");
            var sw = Stopwatch.StartNew();
            lock (_writeLock)
            {
                _transport.WriteFrame(
                    FrameType.Ping,
                    new Dictionary<string, object?> { { "nonce", nonce } });
            }

            Frame reply;
            lock (_readLock)
            {
                reply = ReadFrameWithTimeout(timeoutMs);
            }
            sw.Stop();

            if (reply.Type != FrameType.Pong)
                throw new InvalidOperationException(
                    $"expected PONG reply to PING, got {reply.Type}");
            if (!reply.Meta.TryGetValue("nonce", out var echo) || !Equals(echo, nonce))
                throw new InvalidOperationException(
                    $"PONG nonce mismatch: sent {nonce}, got {echo}");

            return sw.Elapsed;
        }
        finally
        {
            _exchange.Release();
        }
    }

    /// <summary>
    /// Send SHUTDOWN and wait for the child process to exit cleanly.
    /// Returns true if the child exited within <paramref name="timeoutMs"/>;
    /// false means the caller should treat this as an unclean shutdown
    /// (<see cref="Dispose"/> will then force-kill).
    /// </summary>
    public bool Shutdown(int timeoutMs = 5000)
    {
        ThrowIfDisposed();
        if (timeoutMs <= 0) throw new ArgumentOutOfRangeException(nameof(timeoutMs));

        if (!_exchange.Wait(timeoutMs))
            throw new TimeoutException($"another exchange held the kernel for {timeoutMs}ms");
        try
        {
            try
            {
                lock (_writeLock)
                {
                    _transport.WriteFrame(
                        FrameType.Shutdown,
                        new Dictionary<string, object?>());
                }
            }
            catch
            {
                // Pipe already broken — child may have died. Fall through to
                // WaitForExit which will report quickly either way.
            }

            return _process.WaitForExit(timeoutMs);
        }
        finally
        {
            _exchange.Release();
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        // Best-effort polite shutdown if the child is still alive. Skip if
        // someone already called Shutdown() and reaped the process.
        if (!_process.HasExited)
        {
            try
            {
                lock (_writeLock)
                {
                    _transport.WriteFrame(
                        FrameType.Shutdown,
                        new Dictionary<string, object?>());
                }
            }
            catch { /* pipe may already be torn down */ }

            if (!_process.WaitForExit(2000))
                TryKill(_process);
        }

        try { _transport.Dispose(); } catch { /* swallow */ }
        try { _pipe.Dispose(); } catch { /* swallow */ }
        try { _process.Dispose(); } catch { /* swallow */ }
        // _exchange is intentionally not disposed: a Run that's still in
        // its finally block trying to Release() must not see ObjectDisposedException
        // piled on top of the pipe error it's already handling. SemaphoreSlim
        // has no managed resources to leak when AvailableWaitHandle was never
        // accessed (we don't access it).
    }

    // -------------------------------------------------------------------------
    // Internals
    // -------------------------------------------------------------------------

    private static Process StartChild(string pythonExecutable, string pythonPath, string pipeName)
    {
        var psi = new ProcessStartInfo
        {
            FileName = pythonExecutable,
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            RedirectStandardInput = false,
        };

        // netstandard2.0 doesn't expose ProcessStartInfo.ArgumentList, so we
        // assemble the command line ourselves. None of these args (pipe name
        // is a GUID-suffixed identifier) can contain whitespace, so a plain
        // space-join is equivalent to argv-list parsing on both Windows and
        // POSIX. Asserted below so a future caller doesn't slip a path in
        // and silently get tokenised.
        string[] argv = { "-X", "utf8", "-m", "pyexcel.kernel", "--pipe", pipeName };
        foreach (var a in argv)
        {
            if (a.IndexOfAny(new[] { ' ', '\t', '"', '\n' }) >= 0)
                throw new ArgumentException(
                    $"kernel argv entry contains whitespace or quote; would need shell-quoting: {a}");
        }
        psi.Arguments = string.Join(" ", argv);

        // Prepend our pythonPath to whatever PYTHONPATH the user already
        // has so we don't trample legitimate user state.
        var existing = Environment.GetEnvironmentVariable("PYTHONPATH") ?? "";
        psi.Environment["PYTHONPATH"] = existing.Length == 0
            ? pythonPath
            : pythonPath + Path.PathSeparator + existing;

        // Always log unbuffered so any pre-handshake crash hits stderr in time
        // for WaitForConnectOrExit to surface it.
        psi.Environment["PYTHONUNBUFFERED"] = "1";

        var p = new Process { StartInfo = psi, EnableRaisingEvents = true };
        if (!p.Start())
            throw new InvalidOperationException(
                $"failed to start python: {pythonExecutable}");
        return p;
    }

    private static void WaitForConnectOrExit(Task connectTask, Process proc, int timeoutMs)
    {
        var exitWait = Task.Run(() =>
        {
            try { proc.WaitForExit(timeoutMs); } catch { /* ignore */ }
        });

        var winner = Task.WaitAny(new[] { connectTask, exitWait }, timeoutMs);
        if (winner < 0)
            throw new TimeoutException(
                $"kernel did not connect to pipe within {timeoutMs}ms");

        if (proc.HasExited && connectTask.Status != TaskStatus.RanToCompletion)
        {
            // Drain any stderr the child produced — usually a Python traceback
            // explaining why pyexcel.kernel couldn't be loaded.
            var stderr = SafeReadStream(proc.StandardError);
            throw new InvalidOperationException(
                $"kernel exited before connecting (exit={proc.ExitCode}): {stderr.Trim()}");
        }

        // connectTask completed — propagate any pipe-side exception.
        connectTask.GetAwaiter().GetResult();
    }

    private static int Handshake(FrameTransport transport, int timeoutMs)
    {
        // Server sends HELLO first announcing its protocol version. Client
        // replies with HELLO carrying its own. We require an exact match;
        // negotiation can be added when we have >1 supported version.
        transport.WriteFrame(
            FrameType.Hello,
            new Dictionary<string, object?> { { "protocol", (long)Framing.ProtocolVersion } });

        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        var reply = ReadFrameWithDeadline(transport, deadline);

        if (reply.Type != FrameType.Hello)
            throw new InvalidOperationException(
                $"expected HELLO reply during handshake, got {reply.Type}");
        if (!reply.Meta.TryGetValue("protocol", out var pv) || pv is not long pvL)
            throw new InvalidOperationException(
                "HELLO reply missing 'protocol' (long) field");
        if (pvL != Framing.ProtocolVersion)
            throw new InvalidOperationException(
                $"protocol mismatch: server={Framing.ProtocolVersion} client={pvL}");
        return (int)pvL;
    }

    private Frame ReadFrameWithTimeout(int timeoutMs)
    {
        return ReadFrameWithDeadline(_transport, DateTime.UtcNow.AddMilliseconds(timeoutMs));
    }

    internal static Frame ReadFrameWithDeadline(FrameTransport transport, DateTime deadline)
    {
        // ReadFrame is blocking; we approximate a deadline by running it on a
        // worker task and racing against a delay. On timeout the underlying
        // stream is still alive (the read just keeps waiting in the
        // background) — the calling Dispose path will tear it down.
        //
        // Callers from other types in this assembly (or PyExcel.Kernel.Client
        // via InternalsVisibleTo) must hold KernelSupervisor.ReadLock around
        // this call: two concurrent ReadFrame calls on the same pipe corrupt
        // the receive position.
        var task = Task.Run(transport.ReadFrame);
        var remaining = (int)Math.Max(1, (deadline - DateTime.UtcNow).TotalMilliseconds);
        if (!task.Wait(remaining))
            throw new TimeoutException(
                $"no frame received within {remaining}ms");
        return task.Result;
    }

    private static string SafeReadStream(StreamReader r)
    {
        try { return r.ReadToEnd(); } catch { return ""; }
    }

    private static void TryKill(Process? p)
    {
        if (p == null) return;
        try
        {
            if (!p.HasExited) p.Kill();
        }
        catch { /* already dead or insufficient perms */ }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(KernelSupervisor));
    }
}
