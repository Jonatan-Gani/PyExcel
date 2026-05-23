# Phase 4 handoff: the v2 pipe is built — wire it to Excel

> **Audience.** A fresh session picking up Phase 4 ("Excel marshalling
> & first run"). Skim this once; then [`ROADMAP.md`](../ROADMAP.md) is the
> source of truth and the C# / Python tests are the executable spec.

---

## TL;DR

Phase 2 is **complete on both Linux and Windows CI lanes**. C# spawns a
Python kernel, completes the HELLO handshake, sends typed
`RUN_REQUEST` frames carrying Arrow IPC payloads, and gets back
`RUN_RESULT` (or a typed `ERROR`) — over a named pipe on Windows and an
AF_UNIX socket on POSIX.

Phase 4's job is to **wire that into Excel**:

1. The `=PY.RUN(script, range[, kwargs])` UDF, registered with Excel-DNA.
2. C#-side Arrow encoding of `ExcelReference`/`object[,]` range values
   → Arrow IPC bytes that `arrow_io.py` already knows how to decode.
3. Result decoding (Arrow IPC bytes → `object[,]` for spill).
4. A long-lived `KernelSupervisor` + `KernelClient` per workbook,
   owned by `PyExcel.State` (Phase 3) — short-term, can be a
   `static Lazy<KernelClient>` until Phase 3 lands.

Nothing in Phase 4 requires changing the wire protocol or the Python
kernel; all the contracts are pinned.

---

## What's already shipped — public surface

### C# — `PyExcel.Kernel.Client` (typed front-end)

Build a request, hand it to a client, get a result or a typed exception:

```csharp
using var supervisor = KernelSupervisor.StartPython(
    pythonExecutable: @"C:\path\to\python.exe",
    pythonPath:       @"C:\path\to\embedded");  // dir containing pyexcel/

var client = new KernelClient(supervisor);

var result = client.Run(new RunRequest {
    Script    = @"C:\workbook-dir\transform.py",
    Function  = "transform",          // default; can omit
    Arguments = new[] { arrowBytes }, // Arrow IPC streams, one per positional arg
    Kwargs    = new Dictionary<string, object?> { ["factor"] = 5L },
});

// result.Payloads is IReadOnlyList<byte[]>; one Arrow stream, or empty for None.
// result.RunId, result.DurationMs are echoed back.
```

Failures throw `KernelException` with `.Code` (one of `BadRequest`,
`ModuleNotFound`, `ModuleLoadError`, `ModuleExecError`, `FunctionNotFound`,
`FunctionNotCallable`, `BadInput`, `BadReturnType`, `Exception`),
`.PythonType`, `.Message`, `.PythonTraceback`, `.DurationMs`.

Async + cancellation:

```csharp
var result = await client.RunAsync(request, ct);  // OperationCanceledException on ct fire
// client.Cancel(runId) is also exposed for fire-and-forget cancel from
// another thread. The kernel doesn't act on CANCEL yet (see "Known gaps").
```

Streaming events (handlers run synchronously on the Run caller's thread,
between frame reads — keep them cheap):

```csharp
client.ProgressReceived += (s, e) => Trace.WriteLine($"{e.Percent}% {e.Message}");
client.LogReceived      += (s, e) => Trace.WriteLine($"[{e.Level}] {e.Text}");
```

Files: `src/PyExcel.Kernel.Client/*.cs`. Public types: `KernelClient`,
`RunRequest`, `RunResult`, `KernelException`, `ProgressReceivedEventArgs`,
`LogReceivedEventArgs`.

### C# — `PyExcel.Bridge` (transport + lifecycle)

- `KernelSupervisor.StartPython(...)` — spawns the kernel subprocess,
  does the HELLO handshake, returns ready-to-use supervisor.
- `Ping(timeoutMs)` — request/response health check (returns RTT).
- `Shutdown(timeoutMs)` — polite SHUTDOWN frame; reaps the child.
- `Dispose()` — force-kill fallback; **no orphaned python.exe guarantee**.

Internally: dual-lock concurrency model — `ExchangeSemaphore` for
high-level mutual exclusion (one Run/Ping at a time), separate
`WriteLock`/`ReadLock` so `Cancel` can fire while a Run is parked in a
read.

### Python — `embedded/pyexcel/kernel/`

- `arrow_io.py` — `encode(value, orientation=Orientation.COLUMN) -> bytes`
  and `decode(buf) -> Any`. Roundtrips DataFrame / Series / list / tuple /
  numpy 1D-or-2D / scalar. Shape (`table`/`vector`/`scalar`) and vector
  orientation are carried in Arrow schema metadata.
- `worker.py` — pure `run_job(meta, payloads) -> JobOutcome`. The
  supervisor dispatches `RUN_REQUEST` frames through it. Module loading
  is mtime-cached so unchanged scripts don't re-exec.
- `supervisor.py` — the event loop. Already wires `RUN_REQUEST` through
  `worker.run_job` and replies with `RUN_RESULT`/`ERROR`. Phase 4 has
  nothing to add here.
- `transport.py` — POSIX (AF_UNIX) + Windows (`_winapi` against
  `\\.\pipe\<name>`). DACL on the C# side restricts the Windows pipe to
  the current-user SID.

### Wire contract (pinned — don't change)

`RUN_REQUEST` meta:

```json
{
  "run_id":   "<guid>",      // required, echoed back
  "script":   "<abs path>",  // required
  "function": "transform",   // optional, default
  "kwargs":   { ... }        // optional, JSON-serialisable
}
```

`RUN_REQUEST` payloads: zero or more Arrow IPC streams (positional args).

`RUN_RESULT` meta: `{ "run_id", "duration_ms" }`. Payloads:
`[]` for `None` return, `[arrow_bytes]` otherwise.

`ERROR` meta: `{ "run_id", "code", "type", "message", "traceback", "duration_ms" }`.

---

## What Phase 4 needs to build

### 1. Arrow on the C# side — `PyExcel.Excel.ArrowMarshal` (new)

There's currently **no C# Arrow encoder**. The Python side has the full
encoder/decoder; the host needs to:

- Encode `object[,]` / `object[]` (a 2D / 1D range from Excel) →
  Arrow IPC bytes. Shape metadata: `pyexcel-shape = "table"` for 2D,
  `"vector"` for 1D, `"scalar"` for a single cell.
- Decode Arrow IPC bytes (a `RunResult.Payload`) → `object[,]` for
  spill, or a scalar for one-cell result.

Use the `Apache.Arrow` NuGet (supports netstandard2.0 + net48). The
encoded format is `pyarrow.ipc.new_stream(...)` — the C# equivalent is
`Apache.Arrow.Ipc.ArrowStreamWriter` to a `MemoryStream`.

**Don't reinvent the shape conventions.** Read `embedded/pyexcel/kernel/arrow_io.py`'s
module docstring and mirror the schema-metadata keys (`pyexcel-shape`,
`pyexcel-orientation`).

### 2. The `=PY.RUN` UDF — `PyExcel.Excel.PyRunFunction`

Excel-DNA function registration:

```csharp
[ExcelFunction(Name = "PY.RUN", IsThreadSafe = false)]
public static object PyRun(
    string script,
    object input,        // ExcelDna handles object[,] / object / range
    object? kwargs = null)
{
    // 1. Resolve script path (relative → workbook dir; abs → as-is).
    // 2. ArrowMarshal.Encode(input) → byte[]
    // 3. client.Run(new RunRequest { Script = ..., Arguments = [ arrowBytes ], ... })
    // 4. ArrowMarshal.Decode(result.Payload) → object / object[,]
    // 5. Return; Excel spills.
}
```

Threading: Excel UDFs run on a worker pool; KernelClient.Run is
synchronous and serialises through `ExchangeSemaphore`. That's fine for
Phase 4 — one workbook, one kernel, one call at a time. If we want
parallelism later, give each workbook its own kernel.

### 3. Lifecycle — short-term

Until Phase 3's `StateService` lands, use a process-wide
`Lazy<KernelClient>` (with its own KernelSupervisor) so the kernel
boots on the first PY.RUN call and stays warm until Excel exits.
Dispose on add-in unload. See `PyExcel.Addin/Addin.cs` for the
unload hook.

Phase 3 will move ownership into per-workbook state. Phase 4 should
write the dispose logic in a way that's easy to migrate (one method
that hands the supervisor to State.SetKernel(...) or whatever the
final signature looks like).

### 4. Discovery — Python on the user's machine

`KernelSupervisor.StartPython` takes an explicit `pythonExecutable`
path. Phase 4 needs a small `PythonResolver` that finds:

- A venv created by `PyExcel.Setup` (Phase 7) under
  `<workbook-dir>/.pyexcel-venv/`, **or**
- A user-configured python.exe path, **or**
- PATH-discovered `python.exe`.

For first-light Phase 4 just hardcode + document the venv path; the
real resolver is Phase 7 scope.

---

## Known gaps from Phase 2 (not blockers for Phase 4)

- **Worker doesn't emit PROGRESS / LOG yet.** `KernelClient` parses
  them; `worker.py` doesn't produce them. Trivial follow-up — expose a
  `pyexcel.kernel.log(...)` API the user script can call.
- **Worker doesn't handle CANCEL.** `KernelClient.Cancel(runId)` writes
  the frame; the supervisor sees it as "unsupported frame type" and
  replies with an unsolicited ERROR. To make Cancel actually cancel,
  worker needs to run the user function on a separate thread and
  raise `KeyboardInterrupt` when CANCEL arrives. Defer.
- **POSIX pipe file permissions.** On Linux, .NET creates the AF_UNIX
  socket under `/tmp` with default permissions. The Windows side has
  a current-user DACL; the POSIX side relies on the temp-dir guard.
  Multi-user POSIX is not a v2 target, but tightening to mode 0600
  is a small follow-up if needed.
- **`tests/kernel/test_supervisor.py` is Windows-skipped** because the
  Python test plays the C# server role. The C# integration tests
  (`tests/PyExcel.Bridge.Tests/Kernel*Tests.cs`) cover Windows
  end-to-end from the other direction.

---

## Where to look for examples

| What | File |
| --- | --- |
| End-to-end RUN_REQUEST (C# driving) | `tests/PyExcel.Bridge.Tests/KernelClientTests.cs` |
| End-to-end RUN_REQUEST (pytest driving, AF_UNIX) | `tests/kernel/test_supervisor.py` |
| Arrow IPC shape conventions | `embedded/pyexcel/kernel/arrow_io.py` (module docstring) |
| Worker error taxonomy | `embedded/pyexcel/kernel/worker.py` (module docstring) |
| Wire frame layout | `embedded/pyexcel/kernel/framing.py` (module docstring) |
| KernelSupervisor concurrency model | `src/PyExcel.Bridge/KernelSupervisor.cs` class doc |

---

## CI

Both lanes green as of `ae0b4f0`:

- **Cross-platform (netstandard2.0 + kernel tests).** Ubuntu, Python
  3.12, builds netstandard2.0 slice of every project, runs all C#
  tests against the .NET kernel + all 137 pytest tests against the
  Python kernel.
- **Full solution (net48).** Windows, Python 3.12, builds the whole
  solution including net48, runs C# integration tests against the
  Win32 named-pipe transport.

If Phase 4 adds new C# projects, mirror the per-project
`dotnet build --framework netstandard2.0` step in the Linux job
(see `.github/workflows/ci.yml`). The Windows job builds the whole
solution, so it picks them up automatically.
