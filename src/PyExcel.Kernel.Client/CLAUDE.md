# PyExcel.Kernel.Client

## Macro
The typed RPC client over `PyExcel.Bridge.KernelSupervisor`. It builds
`RUN_REQUEST` frames from a typed request, drives the reply loop (progress, log,
result, error), streams progress/log events, and handles cancellation —
including a hard-cancel escalation that kills a wedged kernel so a Run never
waits out its deadline.

## Files
### KernelClient.cs
Typed front-end to the supervisor: executes a job, parses reply frames, raises
progress/log events, and converts a cancelled run into `OperationCanceledException`.
Inputs: a `RunRequest` (script path, function, Arrow-encoded `Arguments`, JSON `Kwargs`)
and an optional timeout / `CancellationToken`. Output: a `RunResult`, or throws
`KernelException` on an ERROR frame / `OperationCanceledException` on cancel; fires
`ProgressReceived` and `LogReceived` events. On cancel it sends a CANCEL frame and, if
the kernel stays unresponsive past the hard-cancel window, kills the child to unblock.

### RunRequest.cs
The request structure for one job. Inputs: `Script` (path), `Function` (defaults to
`transform`), `Arguments` (Arrow IPC byte buffers), `Kwargs` (JSON-serialisable map),
optional caller `RunId`. Output: a value carrying those fields.

### RunResult.cs
A successful job result. Inputs: the echoed `RunId`, kernel `DurationMs`, and zero-or-one
Arrow payload. Output: `RunId`, `DurationMs`, `Payloads`, `IsEmpty`, and a single
`Payload` accessor that throws `InvalidOperationException` when empty.

### KernelException.cs
Exception raised when the kernel returns an ERROR frame. Inputs: run id, error code,
Python exception type name, message, traceback, and duration. Output: an exception with
`RunId`, `Code`, `PythonType`, `PythonTraceback`, and `DurationMs`.

### ProgressReceivedEventArgs.cs
EventArgs for a PROGRESS frame during a run. Inputs: run id, percent (0–100 or null),
message, and the raw frame meta. Output: `RunId`, `Percent`, `Message`, `Meta`.

### LogReceivedEventArgs.cs
EventArgs for a LOG frame during a run. Inputs: run id, level (debug/info/warning/error),
text, and the raw frame meta. Output: `RunId`, `Level`, `Text`, `Meta`.
