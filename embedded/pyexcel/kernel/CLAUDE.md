# kernel

## Macro
The Python kernel: a persistent supervisor + worker that the C# host spawns as
`python -m pyexcel.kernel`. It connects back over a named pipe, performs a HELLO
handshake, then runs user `transform()` functions on demand and marshals results
back. The control plane is length-prefixed binary frames with canonical-JSON
metadata (`framing`); the data plane is Arrow IPC streams carrying shape metadata
(`arrow_io`). This package is the Python counterpart to `src/PyExcel.Bridge` and
`src/PyExcel.Kernel.Client` on the host side.

## Files
### framing.py
Wire-framing codec: length-prefixed binary frames (`body|type|meta_len|meta_json|n_payload|payloads`)
with canonical-JSON metadata. Pure stdlib, size-bounded (256 MiB default), mirrored
byte-for-byte by `PyExcel.Bridge/Framing.cs`. Inputs: a `FrameType`, a meta dict, and
optional payload byte buffers (encode); a `read_exact(n)` callable (decode). Output:
frame `bytes`, or a `Frame(type, meta, payloads)`; raises `FramingError` subclasses
(`FrameTooLargeError`, `TruncatedFrameError`, `MalformedFrameError`). Defines
`PROTOCOL_VERSION = 2`.

### transport.py
Connection layer under `framing`: POSIX `AF_UNIX` and Windows named-pipe client
wrappers that dial into the C#-owned pipe server. Inputs: a pipe name and connect
timeout. Output: a `FrameTransport` exposing `read_exact` / `write_all` / `has_data`;
raises `TransportError` on dial or I/O failure.

### arrow_io.py
Arrow IPC marshalling for the data plane. Inputs: a Python value (`encode` — DataFrame,
list, scalar, or a `types` wrapper) with optional orientation; an Arrow IPC byte buffer
(`decode`). Output: a single Arrow IPC stream `bytes` (with `pyexcel-shape` /
`pyexcel-orientation` schema metadata), or the shape-preserving Python value back
(DataFrame / list / scalar); raises `TypeError` for values Arrow can't represent.

### chart.py
Converts a `transform()` return that is a figure into a typed wire value. Inputs: an
arbitrary value (duck-typed — neither plotly nor matplotlib is imported). Output: a
`types.ChartSpec` (JSON chart-spec document, schema v1) for a Plotly figure, or a
`types.ChartImage` (SVG bytes, PNG fallback) for a Matplotlib figure; raises
`UnsupportedChartTypeError` for an unhandled chart type.

### worker.py
Pure, no-I/O job execution plus the user-facing run helpers. Inputs: a request meta
dict (`run_id`, `script` path, `function`, `kwargs`) and a list of Arrow-encoded
payloads. Output: a `JobOutcome(success, meta, payloads)` — never raises — where a
failure carries an error `code` (`BadRequest`, `ModuleNotFound`, `ModuleExecError`,
`FunctionNotFound`, `BadInput`, `BadReturnType`, `Exception`, `PyExcelInputError`).
Also exposes `is_cancelled()` / `report_progress()` (thread-local per job) and
`install_input_guard()`, which disables `input()` / `sys.stdin` so a console read can't
hang the run.

### supervisor.py
The in-process event loop. Inputs: a pipe name (and connect timeout). Output: the
process exit code; drives the wire by sending `PONG`, `PROGRESS`, and terminal
`RUN_RESULT` / `ERROR` frames. Runs each `RUN_REQUEST` on a worker thread with
cooperative cancellation; a `CANCEL` that a non-cooperative worker ignores is given a
short grace, then the worker is abandoned and an `ERROR`/`Cancelled` reply is sent.

### types.py
User-facing boundary types a `transform()` may accept or return. Inputs: validated
constructor values — `Formula` (A1-mode formula string starting with `=`), `ChartSpec`
(non-blank JSON string), `ChartImage` (image bytes + `svg`/`png` format). Output: the
typed instance; each has a defined Arrow wire representation in `arrow_io`.

### __main__.py
The `python -m pyexcel.kernel` entry point. Inputs: `--pipe <name>` and
`--connect-timeout` argv. Output: exits with the code returned by `supervisor.run`.

### __init__.py
Kernel package surface. Re-exports `ChartImage`, `ChartSpec`, `Formula`,
`is_cancelled`, and `report_progress` for user scripts. Inputs/Output: re-exports only.
