# PyExcel.Excel

## Macro
The Excel integration layer: it turns the ribbon's Run/Import/Export/Paste
intents into kernel exchanges and writes results back into the sheet. It owns the
kernel process lifecycle (`KernelHost`), the run drivers (`RangeRunner` for the
button, `PyRunFunction` for the `=PY.RUN` UDF, `PyRun` as the shared core),
Arrow marshalling between Excel values and the wire, CSV/TSV read/write, and
chart construction. COM- and WinForms-bound services are net48-only; the
marshalling, planners, parsers, and chart-spec logic are cross-platform and
unit-tested on Linux.

## Files
### KernelHost.cs
Process-wide lifecycle wrapper around one `KernelSupervisor` + `KernelClient` pair.
Inputs: none directly — boots from the interpreter and embedded path resolved by
`PythonResolver` against the active workbook. Output: `Client` / `Supervisor` accessors
(boot on first use; a dead kernel is transparently replaced); `Restart()` disposes and
re-arms a fresh boot; `Dispose()` tears the kernel down.

### PyRun.cs
The shared run core: encodes inputs, runs the kernel job, decodes the result, and
archives best-effort. Inputs: a script path, the list of input values, kwargs, a
`KernelClient`, the workbook directory, and optional cancellation + archive context.
Output: the decoded result (DataFrame→2-D table, list→vector, scalar, `ChartSpec`,
`ChartImage`, or `Formula`); throws `KernelException` on a kernel error and re-throws
`OperationCanceledException` on cancel.

### RangeRunner.cs
Drives the ribbon Run button: reads the configured input ranges on Excel's main thread,
runs the kernel exchange on a background task, and writes the result back via
`QueueAsMacro` (net48). Inputs: a `WorkbookState` plus optional progress-factory,
error-display, orientation-chooser, and output-display callbacks. Output: none — side
effects are cells written / charts built, failures surfaced through the error sink, and
(when the selected action opts in) the captured `print()` output shown after a
successful run.

### PyRunFunction.cs
The `=PY.RUN` worksheet function — the Excel-DNA UDF surface over `PyRun`. Inputs: a
script reference and a single argument from the calling cell. Output: the result spilled
into the calling cell, or an Excel error value.

### PythonResolver.cs
Locates the Python interpreter and embedded kernel package for `KernelHost` to spawn.
Inputs: a project directory (nullable). Output: the chosen Python executable path and the
PYTHONPATH directory, resolved in tiers — `PYEXCEL_PYTHON`, a per-workbook
`.pyexcel-venv`, then the Setup-extracted `.pyexcel-kernel` / bundled embedded copy.

### ArrowMarshal.cs
Encodes Excel range values to Arrow IPC streams and decodes kernel results back. Inputs:
`object[,]` / `object[]` / scalar values with optional column names; per-column type
inference (double/bool/string). Output: an Arrow IPC `byte[]`, or a decoded value (2-D
table, 1-D vector, scalar, chart spec, or image).

### ArrowShape.cs
Enumerates the high-level shapes (Table, Vector, Scalar, Chart, Image) and orientation
hints (Row/Column) negotiated through Arrow schema metadata. Inputs: schema metadata
byte strings. Output: `ArrowShape` and `ArrowOrientation` enum values.

### RunProgress.cs
The cross-platform progress abstraction between `RangeRunner` and the WinForms progress
dialog: `IRunProgressSink` (Report / Complete / CancellationToken) plus a `ProgressModel`
of formatting helpers. Inputs: percent (null = indeterminate) and a message. Output:
formatted progress lines and a clamped percent.

### ListOrientation.cs
Enums and `OrientationResolver` describing how a 1-D list result spills into a target
range. Inputs: the target's row and column counts. Output: a `ListOrientation`
(Horizontal/Vertical) and an `OrientationDecision` (resolved, ask-the-user, or invalid).

### SheetSelection.cs
Enum and resolver for how an Excel-import sheet choice resolves. Inputs: an optional
pinned sheet name and the workbook's sheet list. Output: a `SheetResolutionKind`
(Resolved / Prompt) and the selected sheet.

### Formula.cs
Wraps an Excel A1-mode formula string (leading `=`) so the kernel can spill a live
formula instead of a precomputed value. Inputs: formula text. Output: a `Formula`;
validates non-null/non-empty and the leading `=`.

### ChartBuilder.cs
Builds native Excel charts from a validated chart-spec document and embeds rendered
images as worksheet pictures over late-bound COM (net48). Inputs: a `ChartSpecDocument`
or a `ChartImage`, a sheet reference, and position parameters. Output: none — a chart
object or picture on the sheet; cosmetic-step failures go to Trace.

### ChartColor.cs
Parses colour strings (`#RRGGBB`, `#RGB`, `rgb(r,g,b)`, named colours) into OLE
BGR-packed integers. Inputs: a colour string. Output: an `int` RGB value; unrecognised
input defaults to black.

### ChartImage.cs
Holds a kernel-rendered figure image (SVG or PNG bytes). Inputs: image `byte[]` and a
format string. Output: a `ChartImage`; validates non-empty data and a valid format.

### ChartSpec.cs
Wraps a JSON chart-specification string (from a Plotly figure) with structural equality.
Inputs: a JSON string. Output: a `ChartSpec`; validates non-blank JSON.

### ChartSpecDocument.cs
Typed record tree for a parsed chart spec: document root, axes, legend, traces, and
annotations with styling. Inputs: JSON-decoded objects from `ChartSpecParser`. Output:
the strongly-typed `ChartSpecDocument` record tree.

### ChartSpecParser.cs
Parses and validates chart-spec JSON into `ChartSpecDocument`, enforcing version,
supported types, x/y length matching, and unique trace ids. Inputs: a JSON string.
Output: a `ChartSpecDocument`, or a `FormatException` with a user-facing message.

### ChartTraceData.cs
Shapes trace x/y data into COM-ready arrays, handling categorical vs numeric axes,
ISO-8601 dates, and histogram bin labels. Inputs: a `ChartTraceSpec` with x/y lists and a
series type. Output: a shaped record (optional X, Y, and a date flag), or null when no
usable points remain.

### CsvParser.cs
RFC-4180 CSV/TSV parser with Excel extensions (CRLF/LF/CR, BOM stripping, permissive
quoting). Inputs: text or a stream, with optional delimiter and encoding. Output: a list
of records (each a list of field strings); throws `FormatException` on an unterminated
quote.

### CsvWriter.cs
RFC-4180 CSV/TSV writer that round-trips with `CsvParser` (minimal quoting, optional
BOM). Inputs: records of `string?` fields, a delimiter, a line terminator, encoding, and
a BOM flag. Output: a string, or writes to a stream; null fields render empty.

### CsvCellFormatter.cs
Formats an Excel cell value into a CSV field string (invariant round-trip numbers,
`TRUE`/`FALSE`, ISO-8601 dates, null→empty). Inputs: a cell `object?`. Output: a
`string?`.

### CsvCellTypeInference.cs
Infers a cell type from a CSV field string the way Excel import does (invariant double,
boolean tokens; guards leading zeros / plus signs). Inputs: a field string. Output: an
`object?` (null, bool, double, or string).

### ImportPlanner.cs
Pure-logic planner for the Import button: validates fields, resolves the source path,
detects format/delimiter, and parses `path!Sheet`. Inputs: an import-input file path
(optional `!Sheet`), an output range address, and the workbook directory. Output: an
`ImportPlan`; throws `FormatException` on blank/unsupported format.

### ImportService.cs
Drives the Import button (net48): reads CSV/TSV via `CsvParser` or an Excel file via COM,
type-infers CSV cells, and queues the write-back on the main thread. Inputs: a
`WorkbookState` and an optional sheet-chooser callback. Output: none — a table written to
the target range, or an error logged.

### ExportPlanner.cs
Pure-logic planner for the Export button: validates fields, resolves the destination,
detects format from the extension. Inputs: an export-input range, an output file path,
and the workbook directory. Output: an `ExportPlan` (absolute target + delimiter); throws
`FormatException` on blank/unsupported format.

### ExportBatch.cs
Pure validation for the Export Wizard rows: checks each via `ExportPlanner` and detects
duplicate target files. Inputs: a list of (source range, target path) jobs and the
workbook directory. Output: an `ExportBatchValidationResult` (valid flag, error, trimmed
jobs).

### ExportService.cs
Drives the Export button (net48): reads the range on the main thread and writes CSV/TSV
on a background task. Inputs: a `WorkbookState` and an optional batch of export jobs.
Output: none — files written, or an error logged.

### EditIoValidator.cs
Pure validation for the Edit-Import/Export/Paste dialogs, reusing `ImportPlanner` /
`ExportPlanner` as the validity check. Inputs: source/target field strings and the
workbook directory. Output: an `EditIoValidationResult` (valid flag, error, trimmed
values).

### PastePreflight.cs
Pure-logic preflight for Paste: computes the footprint a decoded payload will occupy and
whether the target range already holds content (the destructive-overwrite signal).
Inputs: a decoded value and a target `Value2` snapshot (Excel-DNA sentinels pre-stripped
by the caller). Output: footprint dimensions and a has-content bool.

### PasteService.cs
Drives the ribbon Paste button (net48): parses the OS clipboard as TSV and types each
field into the Destination range. Inputs: a `WorkbookState` (and the clipboard). Output:
none — a table written, or an error logged.

### PastePlanner.cs
Pure-logic planner that selects an archived run's `output.arrow` to paste into a target
range. Currently unused by the ribbon (Paste is a plain clipboard paste) but kept and
unit-tested. Inputs: a target range address and the archived runs. Output: a `PastePlan`;
throws on an invalid target.

### IsExternalInit.cs
Internal compiler polyfill for `System.Runtime.CompilerServices.IsExternalInit` so
records / init-only setters compile on net48 and netstandard2.0. Inputs/Output: none
(compile-time only).
