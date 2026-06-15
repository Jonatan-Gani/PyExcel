# PyExcel.Bridge

## Macro
The host-side bridge: the deterministic framing protocol and the kernel
subprocess lifecycle over a named pipe. This is the C# counterpart to the
Python `embedded/pyexcel/kernel/framing.py` and `transport.py` — the two ends
speak the same wire format. Higher layers (`PyExcel.Kernel.Client`) build typed
requests on top of what this directory exposes.

## Files
### Framing.cs
Wire-protocol codec implementing the frame layout (`body|type|meta_len|meta_json|n_payload|payloads`)
with size-bound enforcement and deterministic encoding. Inputs: a `FrameType`, a meta
dictionary, and optional payload buffers, plus an optional max-frame-bytes cap. Output:
frame `byte[]` (encode) or a `Frame` (decode); throws `FrameTooLargeException` and other
`FramingException` subclasses on protocol violations.

### Frame.cs
Immutable container for one decoded frame. Inputs: a `FrameType`, a meta dictionary, and
optional payload buffers. Output: a `Frame` exposing `Type`, `Meta`, and `Payloads`.

### FrameType.cs
Enumeration of protocol frame types (Hello, Ping, Pong, Shutdown, Error, RunRequest,
RunResult, Progress, Log, Cancel, ListJobs) as stable byte values mirrored from the
Python side. Inputs/Output: byte enum values.

### FrameTransport.cs
Synchronous read/write layer over a `Stream` (named pipe or a test stream) that
encodes/decodes frames. Inputs: a `Stream`, an ownership flag, and an optional max-frame
cap. Output: `Frame` objects via `ReadFrame`; writes via `WriteFrame`; throws
`FramingException` subclasses; disposes the stream when it owns it.

### FramingExceptions.cs
Exception hierarchy for framing failures: `FramingException` (base),
`FrameTooLargeException`, `TruncatedFrameException`, `MalformedFrameException`. Inputs: a
message and optional inner exception. Output: exception instances.

### CanonicalJson.cs
Minimal dependency-free JSON encoder/decoder for frame metadata, matching Python's
`json.dumps` output byte-for-byte. Inputs: object values (bool, numbers, string, dict,
list, null). Output: `byte[]` (encode) or `object?` (decode); throws `FormatException`
on invalid UTF-8 or malformed JSON.

### KernelSupervisor.cs
Owns one Python kernel subprocess: spawns it, completes the HELLO handshake over a
server-side named pipe, and exposes Ping / Shutdown / KillChild plus split read/write
locks so concurrent frame traffic can't interleave. Inputs: a Python executable path,
a PYTHONPATH directory, and optional handshake/frame-size limits. Output: a
`KernelSupervisor` exposing `Process`, `Transport`, the locks, and an `OutputReceived`
event (one event per child stdout/stderr line); throws `TimeoutException` /
`InvalidOperationException` on spawn or handshake failure.

### KernelOutputEventArgs.cs
EventArgs carrying one line of kernel subprocess stdout/stderr (where user `print()`
surfaces). Inputs: an `isError` flag and the line text. Output: `IsError` and `Text`
properties.
