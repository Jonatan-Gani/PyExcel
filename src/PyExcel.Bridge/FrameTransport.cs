using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;

namespace PyExcel.Bridge;

/// <summary>
/// Reads and writes <see cref="Frame"/>s over an arbitrary
/// <see cref="Stream"/>.
///
/// In production the stream is a <see cref="NamedPipeClientStream"/>
/// connected to the Python kernel supervisor (see
/// <see cref="ConnectNamedPipe"/>). Tests inject an in-memory
/// <see cref="MemoryStream"/> or pair the transport against a
/// <see cref="NamedPipeServerStream"/> in the same process.
///
/// This type is intentionally synchronous and blocking. Callers that
/// need to interleave reads and writes from different threads must
/// guard each direction with their own lock — concurrent
/// <see cref="WriteFrame"/> calls on a single pipe will interleave
/// bytes and corrupt the stream.
/// </summary>
public sealed class FrameTransport : IDisposable
{
    private readonly Stream _stream;
    private readonly bool _ownsStream;
    private readonly int _maxFrameBytes;
    private bool _disposed;

    public FrameTransport(
        Stream stream,
        bool ownsStream = true,
        int maxFrameBytes = Framing.DefaultMaxFrameBytes)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _ownsStream = ownsStream;
        _maxFrameBytes = maxFrameBytes;
    }

    /// <summary>
    /// Connect to a named pipe and return a transport that owns the pipe.
    /// Connection is synchronous and will throw <see cref="TimeoutException"/>
    /// if the pipe is not available within <paramref name="connectTimeoutMs"/>.
    /// </summary>
    /// <param name="pipeName">
    /// Pipe name without the <c>\\.\pipe\</c> prefix. On Linux (.NET 5+ test
    /// hosts) this maps to a Unix-domain-socket-backed pipe.
    /// </param>
    public static FrameTransport ConnectNamedPipe(
        string pipeName,
        int connectTimeoutMs = 5000,
        int maxFrameBytes = Framing.DefaultMaxFrameBytes)
    {
        if (string.IsNullOrWhiteSpace(pipeName))
            throw new ArgumentException("pipe name must be non-empty", nameof(pipeName));

        // Local pipe (server is "." == this machine). PipeOptions.Asynchronous
        // is enabled even for synchronous use — without it overlapped I/O is
        // disabled and concurrent reads/writes on the same pipe deadlock.
        var pipe = new NamedPipeClientStream(
            serverName: ".",
            pipeName: pipeName,
            direction: PipeDirection.InOut,
            options: PipeOptions.Asynchronous);

        try
        {
            pipe.Connect(connectTimeoutMs);
        }
        catch
        {
            pipe.Dispose();
            throw;
        }

        return new FrameTransport(pipe, ownsStream: true, maxFrameBytes: maxFrameBytes);
    }

    /// <summary>Read the next frame from the stream. Blocks until a full
    /// frame is available or EOF is hit mid-frame.</summary>
    public Frame ReadFrame()
    {
        ThrowIfDisposed();
        return Framing.ReadFrame(_stream, _maxFrameBytes);
    }

    /// <summary>Encode and write one frame to the stream. Calls
    /// <see cref="Stream.Flush"/> so the bytes reach the peer immediately.</summary>
    public void WriteFrame(
        FrameType type,
        IReadOnlyDictionary<string, object?> meta,
        IReadOnlyList<byte[]>? payloads = null)
    {
        ThrowIfDisposed();
        var wire = Framing.EncodeFrame(type, meta, payloads, _maxFrameBytes);
        _stream.Write(wire, 0, wire.Length);
        _stream.Flush();
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        if (_ownsStream) _stream.Dispose();
    }

    private void ThrowIfDisposed()
    {
        if (_disposed) throw new ObjectDisposedException(nameof(FrameTransport));
    }
}
