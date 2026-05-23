using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PyExcel.Bridge;
using Xunit;

namespace PyExcel.Bridge.Tests;

public class FrameTransportTests
{
    // -------------------------------------------------------------------------
    // MemoryStream-backed roundtrip
    // -------------------------------------------------------------------------

    [Fact]
    public void Write_Then_Read_RoundTrips()
    {
        using var ms = new MemoryStream();
        using (var writer = new FrameTransport(ms, ownsStream: false))
        {
            writer.WriteFrame(
                FrameType.Hello,
                new Dictionary<string, object?> { { "protocol", (long)Framing.ProtocolVersion } });
            writer.WriteFrame(
                FrameType.RunResult,
                new Dictionary<string, object?> { { "status", "done" } },
                new[] { Encoding.UTF8.GetBytes("payload-bytes") });
        }

        ms.Position = 0;
        using var reader = new FrameTransport(ms, ownsStream: false);
        var hello = reader.ReadFrame();
        var result = reader.ReadFrame();

        Assert.Equal(FrameType.Hello, hello.Type);
        Assert.Equal((long)Framing.ProtocolVersion, hello.Meta["protocol"]);
        Assert.Empty(hello.Payloads);

        Assert.Equal(FrameType.RunResult, result.Type);
        Assert.Equal("done", result.Meta["status"]);
        Assert.Single(result.Payloads);
        Assert.Equal(Encoding.UTF8.GetBytes("payload-bytes"), result.Payloads[0]);
    }

    [Fact]
    public void ReadFrame_OnEmptyStream_Throws()
    {
        using var ms = new MemoryStream(Array.Empty<byte>());
        using var transport = new FrameTransport(ms, ownsStream: false);
        Assert.Throws<TruncatedFrameException>(() => transport.ReadFrame());
    }

    [Fact]
    public void Dispose_ClosesOwnedStream()
    {
        var ms = new MemoryStream();
        var transport = new FrameTransport(ms, ownsStream: true);
        transport.Dispose();
        // Disposed MemoryStream throws on access.
        Assert.Throws<ObjectDisposedException>(() => ms.WriteByte(0));
    }

    [Fact]
    public void Dispose_LeavesNonOwnedStreamOpen()
    {
        using var ms = new MemoryStream();
        var transport = new FrameTransport(ms, ownsStream: false);
        transport.Dispose();
        // Still writable.
        ms.WriteByte(0);
    }

    [Fact]
    public void Operations_AfterDispose_Throw()
    {
        using var ms = new MemoryStream();
        var transport = new FrameTransport(ms);
        transport.Dispose();
        Assert.Throws<ObjectDisposedException>(() => transport.ReadFrame());
        Assert.Throws<ObjectDisposedException>(() =>
            transport.WriteFrame(FrameType.Ping, new Dictionary<string, object?>()));
    }

    [Fact]
    public void EncodeFrame_RespectsMaxFrameBytes()
    {
        using var ms = new MemoryStream();
        using var transport = new FrameTransport(ms, ownsStream: false, maxFrameBytes: 64);
        var tooBig = new byte[100];
        Assert.Throws<FrameTooLargeException>(() =>
            transport.WriteFrame(
                FrameType.RunResult,
                new Dictionary<string, object?>(),
                new[] { tooBig }));
    }

    // -------------------------------------------------------------------------
    // Named-pipe roundtrip — pairs ConnectNamedPipe against a server we host
    // in-process. Validates the wire format survives a real OS IPC channel
    // (Windows named pipe / Linux Unix-domain socket on .NET 5+).
    // -------------------------------------------------------------------------

    [Fact]
    public async Task NamedPipe_Bidirectional_Roundtrip()
    {
        var pipeName = "pyexcel-test-" + Guid.NewGuid().ToString("N");

        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            transmissionMode: PipeTransmissionMode.Byte,
            options: PipeOptions.Asynchronous);

        var serverTask = Task.Run(async () =>
        {
            await server.WaitForConnectionAsync();
            // Server reads the client's HELLO, replies with PONG.
            var hello = Framing.ReadFrame(server);
            Assert.Equal(FrameType.Hello, hello.Type);
            Assert.Equal("client", hello.Meta["from"]);

            var reply = Framing.EncodeFrame(
                FrameType.Pong,
                new Dictionary<string, object?> { { "echo", hello.Meta["from"] } });
            server.Write(reply, 0, reply.Length);
            server.Flush();
        });

        using (var client = FrameTransport.ConnectNamedPipe(pipeName, connectTimeoutMs: 5000))
        {
            client.WriteFrame(
                FrameType.Hello,
                new Dictionary<string, object?> { { "from", "client" } });

            var pong = client.ReadFrame();
            Assert.Equal(FrameType.Pong, pong.Type);
            Assert.Equal("client", pong.Meta["echo"]);
        }

        await serverTask;
    }

    [Fact]
    public void NamedPipe_Connect_ToNonexistent_Throws()
    {
        // No server listening on this name — expect TimeoutException.
        var pipeName = "pyexcel-test-noserver-" + Guid.NewGuid().ToString("N");
        Assert.Throws<TimeoutException>(() =>
            FrameTransport.ConnectNamedPipe(pipeName, connectTimeoutMs: 200));
    }

    [Fact]
    public async Task NamedPipe_ServerDisconnect_TriggersTruncated()
    {
        var pipeName = "pyexcel-test-disconnect-" + Guid.NewGuid().ToString("N");

        using var server = new NamedPipeServerStream(
            pipeName,
            PipeDirection.InOut,
            maxNumberOfServerInstances: 1,
            transmissionMode: PipeTransmissionMode.Byte,
            options: PipeOptions.Asynchronous);

        var serverTask = Task.Run(async () =>
        {
            await server.WaitForConnectionAsync();
            // Drop the connection without sending anything.
            server.Disconnect();
        });

        using var client = FrameTransport.ConnectNamedPipe(pipeName, connectTimeoutMs: 5000);
        await serverTask;

        // Reading from a closed peer raises TruncatedFrameException on the
        // first read (no length prefix arrived).
        Assert.Throws<TruncatedFrameException>(() => client.ReadFrame());
    }

    // -------------------------------------------------------------------------
    // Multiple frames back-to-back through one transport
    // -------------------------------------------------------------------------

    [Fact]
    public void StreamingFrames_BackToBack()
    {
        using var ms = new MemoryStream();
        using (var writer = new FrameTransport(ms, ownsStream: false))
        {
            for (var i = 0; i < 5; i++)
                writer.WriteFrame(
                    FrameType.Progress,
                    new Dictionary<string, object?> { { "pct", (long)(i * 20) } });
        }

        ms.Position = 0;
        using var reader = new FrameTransport(ms, ownsStream: false);
        var pcts = Enumerable.Range(0, 5)
            .Select(_ => (long)reader.ReadFrame().Meta["pct"]!)
            .ToArray();
        Assert.Equal(new long[] { 0, 20, 40, 60, 80 }, pcts);
    }
}
