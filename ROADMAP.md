# PyExcel — Roadmap to Production

> **Purpose of this file.** The complete, ordered task list to take PyExcel from
> its current state (a debt-laden v1 VBA add-in + a 5%-built v2 .NET skeleton) to
> a single production-grade v2 .NET product. For *where the code lives*, see
> [`ARCHITECTURE.md`](ARCHITECTURE.md).
>
> **How to use it.** Phases are sequential and dependency-ordered. Each phase has
> an objective, deliverables, a task checklist, and exit criteria. A phase is not
> "done" until it also meets the *Definition of production-grade* below.

## Goal & end state

Move everything possible out of v1 VBA into v2 .NET. The Python side shrinks to a
single kernel package (it must stay Python — it runs user `transform()` code).

| Layer | Today | Done |
| --- | --- | --- |
| Orchestration, Excel I/O, UI, setup | ~12k lines VBA | **0 lines VBA** — all .NET |
| Excel ↔ Python IPC | XML files + `meta.xml` polling | named pipe + framing + Arrow |
| Python | `tools.py` + `xmlParsing.py` (~3.7k lines, ~60% dead) | one `pyexcel.kernel` package |
| Distribution | `PyExcel.xlam` | `PyExcel-AddIn64.xll` |

## Status & progress log

**Current position:** Phase 0 ✅ · Phase 1 ✅ · Phase 2 ✅ · Phase 4 mostly done (everything except items that hinge on Phase 3's COM-bound pieces) · Phase 3 foundation landed (`StateService`, ribbon wiring, `ScriptDirectoryWatcher`) **plus the `AppEventSink` + CustomXMLPart persistence COM glue (code-complete, awaiting Windows smoke test)**. The #4 UDF cancel bridge just landed (`ExcelAsyncUtil.Observe` + cancel-on-dispose → `KernelClient.RunAsync` token); only #5's progress WinForm is left on Phase 4's polish list — see [`docs/phase3-and-4-completion.md`](docs/phase3-and-4-completion.md). Working branch: `NET_Migration`.

Terse running record (newest first) so a new session can pick up where work stopped:

- **2026-05-30 — Phase 4: archive a run — `RunArchive` + retention cap.** Closed the last open cross-platform Phase 4 checkbox. New `PyExcel.State.RunArchive` is the on-disk diagnostic store: each call to `Archive(RunArchiveEntry)` writes a directory `{root}/{yyyyMMddTHHmmssfff}_{8-hex-rand}/` containing the encoded inputs (`input_0.arrow`, `input_1.arrow`, …, Arrow IPC bytes, one per positional argument), the output (`output.arrow`, present iff the run returned a non-`None` payload), the formatted error block (`error.txt`, the same `KernelErrorRecord.FormatForClipboard()` the ribbon's Copy Last Error button uses — present on `Error`/`Cancelled`), and a line-per-field manifest (`manifest.txt`: `RunId`, `TimestampUtc`, `DurationMs`, `Source`, `WorkbookKey` (omitted when unbound), `ScriptPath`, `Function`, `InputCount`, `Status`, plus `ErrorCode`/`ErrorType`/`ErrorMessage` on failures — multi-line error messages collapsed to a single line so the file stays parseable; the full traceback lives in `error.txt` next door). Retention cap defaults to 20 (FIFO eviction by directory-name lex order, which matches chronological because the run-id leads with the UTC timestamp); pruning runs synchronously at the tail of every `Archive` call under the service's single coarse lock so concurrent writers can't trip each other's eviction. New `RunArchiveEntry` (record carrying timestamp + workbook key + script + function + source + duration + inputs + output + error + status), new `ArchivedRun` (record returned by `List()` — parsed headline fields plus the directory path), new `RunArchiveContext` (the per-callsite wrapper threading the archive instance + a `Source` label + the current workbook key through), new `RunArchiveStatus` enum (`Success` / `Error` / `Cancelled` — split out because the kernel reports cancel distinctly and a replayer wants to know "did the user pull the rug" vs. "did the script blow up"). Wire-in: `PyRun.Execute` / `ExecuteMany` / `ExecuteAsync` / `ExecuteManyAsync` gain an optional `RunArchiveContext? archive = null` trailing parameter; when supplied, a `Stopwatch`-timed try/finally captures the kernel exchange and `ArchiveBestEffort` writes a `RunArchiveEntry` (success, kernel error with `KernelException` mapped to `Cancelled`-vs-`Error` by code, host error, or `OperationCanceledException` → `Cancelled`). The archive write is wrapped in its own try/catch so an I/O fault in the side path can't mask the user-facing result or the in-flight exception. `PyRunFunction.RunCoreAsync` and `RangeRunner.RunActiveScript` build their archive contexts from `PyExcelServices.RunArchive` + the active workbook key, tagging Source `"PY.RUN"` and `"Run Python button"` respectively. New `PyExcelServices.RunArchive` static slot defaults to a `RunArchive` rooted at `%LOCALAPPDATA%\PyExcel\runs` (XDG-derived on Linux, `Path.GetTempPath()` fallback when neither is set) — lazy directory creation so the default service doesn't litter the FS on a fast-path exit. Coverage: 22 `RunArchiveTests` (construction validation, success/error layout, manifest field-by-field, multi-line error message collapse, retention prune, `MaxRuns=0` edge, `List()` ordering + manifest parse, corrupt-manifest skip, 48-write concurrent prune cap); 3 e2e `PyRunTests` through a real kernel (`Execute_WithArchiveContext_ArchivesSuccessfulRun`, `…_ArchivesKernelError`, `…_NoArchiveContext_DoesNotArchive`).
- **2026-05-30 — Phase 4: better error UI — `ErrorService` + ribbon Show/Copy Last Error.** Promoted the "Surface kernel errors" bullet from the partial state (LogDisplay only) toward done by adding the ribbon-side surface the partial entry called out as missing. New `PyExcel.State.ErrorService` keeps the last error per workbook key plus a global fallback slot for failures that arrive without a bound workbook (kernel boot, add-in init). New `KernelErrorRecord` is the immutable carrier — timestamp, source label, kernel `Code` / `PythonType`, `Message`, full `PythonTraceback`, optional `ScriptPath` — with `FormatForClipboard()` rendering a stable multi-line block ready to paste into a bug report. `PyRunFunction.RunCoreAsync` and `RangeRunner.RunActiveScript` route their existing `KernelException` / `Exception` catches through the service (best-effort, swallowed if the registry throws — the cell-level `#VALUE!` is never blocked by error-recording). Ribbon gains an Errors group (`grpErrors`) with **Show Last Error** (writes the formatted block to Excel-DNA `LogDisplay` and `Show()`s the window) and **Copy Last Error** (`System.Windows.Forms.Clipboard.SetText` of the formatted block — ribbon callbacks already run on Excel's STA main thread). `getEnabled` is `RibbonHasError`, independent of the `RibbonEnabled` workbook-level toggle, so errors surface even when PyExcel is disabled. `PyExcel.Ribbon.csproj` gains an explicit `<Reference Include="System.Windows.Forms" />` (SDK-style net48 doesn't auto-reference it). Coverage: 11 `ErrorServiceTests` (per-workbook scoping, global fallback, override semantics, replace-on-record, clear-fires-event-even-if-empty, format-includes-every-field, format-omits-redundant-type, format-omits-empty-optionals, null-record-throws). Still outstanding on this bullet: per-script error pane + cell hover tooltip (both Phase 8 WinForms).
- **2026-05-30 — Phase 4: `ExcelFormula` A1-mode round-trip (marshalling layer).** Closed the Phase 4 "`ExcelFormula` round-trip" bullet. New typed wrapper on both sides — Python `pyexcel.kernel.Formula(text)` (frozen dataclass, validates leading `=`) and C# `PyExcel.Excel.Formula(text)` (sealed class with ordinal equality) — backed by an Arrow string column whose field carries `pyexcel-cell-type=formula` metadata. Encoders detect the type: pyarrow's `_value_to_table` routes a scalar `Formula` and an all-`Formula` list/tuple through string Arrow with the marked field (mixed `Formula`/non-`Formula` lists are rejected — the marker is per-column, so there's no clean wire representation). C# `ArrowMarshal.ScanTypes` gains a `canFormula` arm + `BuildColumn` emits the formula metadata when every non-null cell is a `Formula`. Decoders thread the schema through `DecodeTable` / `DecodeVector` / `DecodeScalar`: any field with the marker wraps cells in `Formula(text)` rather than returning the raw string. Python's `_table_to_python` handles mixed-column tables (some formula columns, some regular) by stitching per-column Series so non-formula columns keep their Arrow-inferred dtype. **Out of scope:** DataFrame-of-Formula encoding (skipped — `pa.Table.from_pandas` flattens object Series to plain strings; the host would need to construct the Arrow table by hand) and the actual `Range.Formula = …` write hook in `RangeRunner` (the marshalling round-trip is what this item delivered; the COM-bound write is a separate ribbon-integration follow-up). Tests: 7 C# unit tests cover scalar, vector with nulls, mixed-with-double-column table, the mixed-within-column string fallback, plus Formula equality / validation; 8 Python tests cover scalar / vector / mixed-list-rejection / DataFrame-skipped / dataclass invariants. 2 end-to-end PyRun tests pin the cross-language round-trip through a real kernel — `from pyexcel.kernel import Formula; def transform(): return Formula('=SUM(A1:B2)')` decodes as a C# `Formula("=SUM(A1:B2)")`. Python lane verified locally (152/152 + 1 intentional skip); C# side awaits CI.
- **2026-05-30 — Phase 4: Range↔Arrow date/null/formula round-trip tests + date-decode bug fix.** Closed the Phase 4 "tests for dates, nulls, formulas" bullet, and surfaced a real bug while doing it. `ArrowMarshal.ReadCell` was falling through to its last-ditch `_ => array.GetType().Name` arm for `Date32Array` / `Date64Array` / `TimestampArray`, so a Python script returning a `datetime.datetime` would render the literal string `"TimestampArray"` in the user's cell. The new decoder handles all four `TimeUnit`s for `TimestampArray` plus `Date32` / `Date64`, landing on a naive `DateTime` (`DateTimeKind.Unspecified`) — pyarrow's default for a Python `datetime` is `timestamp[us]` with no timezone, and that's every path PyExcel currently exercises (matching pandas's `datetime64[ns]` → `timestamp[ns]` via the nanosecond arm). New tests: 7 unit tests in `ArrowMarshalTests` build `Date32` / `Date64` / `TimestampArray` buffers directly via Apache.Arrow's builders (the C# encoder doesn't produce these types, so coverage is decode-only) and assert the right `DateTime` falls out; 4 end-to-end tests in `PyRunTests` pin the cross-language round-trip via a real kernel — `datetime.datetime` → `timestamp[us]`, `datetime.date` → `date32`, list-of-datetime → vector spill, `pd.to_datetime` series → `timestamp[ns]`. Null coverage gained 3 unit tests (nulls embedded in numeric / bool / vector columns survive C#-only round-trip) and 2 e2e tests covering each leg in isolation (`pandas.df.isna().sum()` proves C# nulls reach the kernel as missing; `[1.0, None, 3.0]` from Python proves kernel nulls decode to C# null) — the full-table round-trip is fragile because `to_pandas` converts Arrow nulls to NaN, so we test each direction instead. Formula-result coverage adds 3 regression tests pinning that `=SUM`-style doubles and `=TEXT`-style strings round-trip via the per-column type inference and that mixed-type columns (a `#DIV/0!` string sitting in a numeric column) fall back to string without throwing. Single-locale assumption noted: the string-fallback path uses `double.ToString()` which is locale-dependent; the formula tests use whole numbers (`1`, `3`) that stringify identically across locales.
- **2026-05-30 — Phase 4 #4: UDF → kernel cancel bridge.** Closed the open polish item from `docs/phase3-and-4-completion.md` §4. `PyRunFunction.Run` now routes through `ExcelAsyncUtil.Observe` instead of `ExcelAsyncUtil.Run`: a small `IExcelObservable` (`PyRunObservable`) starts the work on the threadpool with a `CancellationTokenSource`, and returns a `CancelOnDispose(cts)` to Excel-DNA. When Excel-DNA disposes the subscription (formula change, workbook close, cell delete), the CTS is cancelled; the token registration inside `KernelClient.RunAsync` pushes a `CANCEL` frame to the kernel, which sets `is_cancelled()` and replies `ERROR/Cancelled`. The handler swallows that as a no-op (the observer is already gone). `KernelException` and other faults still surface as `#VALUE!` + LogDisplay+Trace, matching the prior contract. New async overloads `PyRun.ExecuteAsync` / `ExecuteManyAsync` keep the cancellation token plumbed through the marshal-and-dispatch core; the sync `Execute` / `ExecuteMany` are unchanged so `RangeRunner` (the ribbon button) stays put. Tests: `RunAsync_Cancellation_Surfaces_OperationCanceledException` covers the supervisor-pumps-CANCEL path at the `KernelClient` level; `ExecuteAsync_TokenCancelledMidRun_ThrowsOperationCanceled` covers the same end-to-end through `PyRun`; happy-path tests pin that the async wrappers preserve sync semantics. The Excel-DNA observable itself (`PyRunObservable`) is `#if NETFRAMEWORK` and isn't unit-testable without Excel-DNA — its CI coverage is "the net48 build still compiles with the new wiring". Phase 4's polish list now has only #5's WinForms progress dialog left (Phase 8).
- **2026-05-29 — Phase 3 #2 `AppEventSink` + #3 COM-side CustomXMLPart persistence.** Closed the two COM-coupled Phase 3 tails. New `AppEventSink` (`PyExcel.Addin`, net48) subscribes to `Application.{WorkbookOpen, WorkbookActivate, WorkbookBeforeSave, WorkbookBeforeClose, SheetActivate}`: Open restores state from the workbook's `CustomXMLPart` + repaints the ribbon; Activate repaints (the active key changed but nothing in the registry mutated, so no `StateChanged` fires on its own); BeforeSave flushes state into the part; BeforeClose flushes + `StateService.Forget`; SheetActivate → `SetCurrentSheet`. Every handler runs inside a `Guard` so a fault can't propagate into Excel's event pump. New `WorkbookStatePersister` (net48) is the thin COM shell over `CustomXMLParts.SelectByNamespace`/`Add`/`Delete`, delegating the XML round-trip to the already-tested `WorkbookStateCodec` (new pure helpers `SerializeToString` + `TryDeserialize`, the latter returning `false` on null/foreign/corrupt/wrong-version XML so a bad part can never throw into a COM handler). **Design fork resolved (PIA vs. late-bound):** app events can't be wired over a late-bound `dynamic` (`+=` needs a typed delegate), so this is the one place PyExcel takes a typed Office interop reference — via the `ExcelDna.Interop` NuGet package — restorable on the Excel-less Windows CI runner (unlike an *installed* PIA), with its own MSBuild `.targets` wiring up the references. That dissolves the doc's only objection to the typed route; everything else COM-bound (`ExcelWorkbookContext`, `RangeRunner`) stays late-bound. *(The package's implicit embed-via-`.targets` mechanism didn't surface the Office types on the Windows CI runner — CS0234 — so the PIA assemblies it restores (`lib/net452`) are referenced explicitly with `EmbedInteropTypes` instead; see `PyExcel.Addin.csproj`.)* `WorkbookKeys` now owns the shared session GUID + saved-vs-unsaved key rule so the sink and `ExcelWorkbookContext` derive **identical** keys (required for correctness — otherwise per-workbook state / `Forget` / persist would target different entries); `ExcelWorkbookContext` refactored onto it. Ribbon gained `PyExcelServices.RequestRibbonInvalidate` (registered in `RibbonOnLoad`) so the sink can repaint on Activate without reaching into the ribbon class. **Verification:** the cross-platform logic is unit-tested — `WorkbookKeysTests` (8), `WorkbookStatePersistenceTests` (9), `PyExcelServicesTests` (1) — and builds on both CI lanes; the COM plumbing (`AppEventSink`, persister) is compiled by the Windows CI lane (full-solution build, NuGet-restored PIA) but its **runtime behaviour still needs the Windows + Excel smoke test** in `docs/phase3-and-4-completion.md` §2/§3 — *not run yet* (this was a Linux session with no dotnet/Excel). Closes Phase 3's "state survives save/reopen" exit criterion pending that smoke test.
- **2026-05-29 — Phase 4/#5 kernel half: `report_progress` → `PROGRESS` frames.** Closed the Phase 2 "worker doesn't emit PROGRESS" gap. New user-facing `pyexcel.kernel.report_progress(percent=None, message="")` (re-exported from `pyexcel.kernel` next to `is_cancelled`): a long-running `transform()` calls it to push status to the host. Mirrors the cancellation plumbing — `worker._begin_job(event, progress_sink)` installs a per-job sink, `report_progress` forwards `(float|None, str)` to it (inert no-op when no job is in flight), `_end_job` clears it. The supervisor's `_run_with_cancellation` owns a `queue.Queue` the sink enqueues onto from the worker thread; the main loop drains it to `PROGRESS` frames via a new `_flush_progress` helper — every wire write stays on the main thread (single writer, no race with PONG/terminal), flushed each ~50 ms tick and once more after the worker joins so all `PROGRESS` precede the terminal `RUN_RESULT`/`ERROR`. Meta matches `KernelClient.RaiseProgress`: `run_id`, `percent` (JSON `null` = indeterminate), `message`. Pure Python — no C# touched; `FrameType.Progress` and `KernelClient.ProgressReceived` were already wired. 5 worker unit tests + 1 e2e supervisor test; **146/146 kernel tests green**. Remaining for #5 is the WinForms progress dialog (Windows-only, Phase 8); the C#-only open items (#2 `AppEventSink`, #3b CustomXMLPart persister, #4 UDF cancel bridge) are untouched — they need a Windows + .NET SDK box this Linux session doesn't have.
- **2026-05-29 — Phase 4 VERIFIED on real Excel + the packaging fixes that took to get there.** Smoke-tested the whole COM path end-to-end on Windows + 64-bit Excel: Enable PyExcel → type Script/Input/Output → Run → result spills back. All three cases pass — single-table doubling (`A1:B2`→`D1:E2`), error surfacing (`raise ValueError` shows the full Python traceback in Excel-DNA LogDisplay, sheet untouched), and multi-input (`prices=A1:A2; signals=B1:B2` → two positional args). This closes the Phase 4 exit criterion. Five fixes were needed that CI could not have caught (no Excel on the runners):
  - **`Microsoft.CSharp` reference** on the net48 targets of `PyExcel.Addin` + `PyExcel.Excel` — `dynamic` COM access needs the runtime binder, which SDK-style net48 projects don't reference by default (CS0656). *This one would have failed the Windows CI build too.*
  - **`.dna` `<ExternalLibrary>` entries for `PyExcel.Ribbon.dll` + `PyExcel.Excel.dll`** — Excel-DNA only scans listed libraries for `[ExcelRibbon]` / `[ExcelFunction]`; a merely-referenced dependency is packed but not scanned, so the ribbon tab never rendered (AutoOpen ran but `RibbonOnLoad` never fired) and `=PY.RUN` wasn't registered.
  - **`OnEnablePyExcel` wired to toggle `Enabled`** — was a Phase-1 stub, so every `getEnabled=RibbonEnabled` control stayed greyed out with no way to turn the add-in on.
  - **`app.config` binding redirects** for Apache.Arrow's `System.*` deps (Unsafe 4.0.6.0, Memory 4.0.1.2, Buffers 4.0.3.0, Numerics.Vectors 4.1.4.0, Tasks.Extensions 4.2.0.1) — net48 doesn't auto-unify like .NET Core, so the first `ArrowMarshal` call threw `0x80131040`. ExcelDna copies app.config → `PyExcel-AddIn64.xll.config`.
  - Confirmed `ExcelWorkbookContext` (#1) returns sane per-workbook keys in the live host. #1 and #7 are now verified, not just code-complete.
- **2026-05-25 — Phase 4 step 8: `OnRunPython` ribbon button (Phase 4 exit criterion).** New `PyExcel.Excel.RangeRunner` (net48-only, `#if NETFRAMEWORK`) drives the button end-to-end: on the main thread it reads the configured input ranges off the active sheet (re-basing Excel's 1-based `Value2` arrays to 0-based `object?[,]`, scalars pass through), then dispatches `PyRun.ExecuteMany` on a background `Task` (SAFE-1: the callback returns before any pipe traffic), then writes the decoded result back into the `PyOutput` range via `ExcelAsyncUtil.QueueAsMacro` (a table resizes the anchor to its footprint, a scalar drops into the top-left, a `None` return writes nothing). All COM is late-bound `dynamic` on `ExcelDnaUtil.Application` — no Office PIA, so the PIA-less Windows CI build still compiles. `PyExcelRibbon.OnRunPython` now delegates to it (PyExcel.Ribbon gains a project ref to PyExcel.Excel; PyExcel.Excel gains one to PyExcel.State for `RibbonRangeParser` + `WorkbookState`). Errors surface to `Trace` + Excel-DNA `LogDisplay`, same contract as the `=PY.RUN` UDF. Can't be CI-tested (COM range I/O); the step-by-step Windows smoke test is in `docs/phase3-and-4-completion.md`. The remaining open items are #2 `AppEventSink` (blocked on a PIA-vs-late-bound design decision — the PIA breaks CI), #4 cancel bridge, #5 progress UI.
- **2026-05-25 — Phase 3 step: `ExcelWorkbookContext`.** Concrete `IWorkbookContext` backed by `Application.ActiveWorkbook` over the Excel COM interop. Uses `ExcelDnaUtil.Application` dispatched through `dynamic` so `PyExcel.Addin` doesn't need an Office PIA reference. Key strategy: saved workbooks → `Workbook.FullName`; unsaved (empty `Path`) → `"unsaved:{SessionGuid}:{Workbook.Name}"` where `SessionGuid` is allocated once per add-in load so `Book1` / `Book2` don't collide. Any COM exception during lookup yields `null` (the ribbon's getters all tolerate that as "no workbook"). Wired into `PyExcelServices.WorkbookContext` from `AddIn.AutoOpen`. Can't be CI-tested (no Excel COM on Linux); manual smoke test on a Windows + Excel dev box is the verification path tracked in `docs/phase3-and-4-completion.md` §1.
- **2026-05-25 — Phase 3 step: `WorkbookStateCodec`.** Pure-XML round-trip for the user-editable slice of `WorkbookState`. Schema is hand-rolled `XDocument`: root `<pyexcel state-version="1" xmlns="urn:pyexcel:state:1">` carrying `<enabled>`, optional `<selected-script>` / `<py-input>` / `<py-output>` / `<selected-action>`, and an `<actions>` list whose elements have `name`/`script`/`input`/`output` attributes plus an optional `<kwargs>` child sorted by key for byte-stable output. Persisted fields are the ones the user types into the ribbon; transient ones (`CurrentSheet`, `AvailableScripts`, `WorkbookKey`) stay out of the XML — the caller supplies the key on Deserialize so a workbook-saved-as-copy gets the right new key, and the live sources (sheet-activate event, `ScriptDirectoryWatcher`) repopulate the transients. Errors throw `FormatException`: wrong root, wrong namespace, missing/unsupported `state-version`, missing required `<action>` attribute. 15 round-trip + error-path tests in `WorkbookStateCodecTests.cs`. The COM-side persister (`CustomXMLPart` read/write on `WorkbookOpen` / `WorkbookBeforeSave`) is the Windows-only follow-up tracked in `docs/phase3-and-4-completion.md` §3.
- **2026-05-25 — Phase 4 step 7: `RibbonRangeParser`.** Pure-logic parser for the ribbon's Input / Output text fields. Accepts the `prices=A1:C10; signals=D1:D10` multi-binding syntax plus the existing anonymous `A1:C10` form; returns an ordered list of `RangeBinding(Name?, RangeText)` records. Whitespace tolerance around separators and `=`; empty entries (trailing/leading/double `;`) silently skipped. Malformed input throws `FormatException`: empty name before `=`, empty range after `=`, or duplicate name (case-sensitive — matches the rest of the codebase). 15 unit tests in `RibbonRangeParserTests.cs`. Unblocks #7 (`OnRunPython` ribbon button) per `docs/phase3-and-4-completion.md`.
- **2026-05-24 — Phase 3 foundation: state + ribbon wiring + script watcher.** New `PyExcel.State` assembly (multi-targets net48 + netstandard2.0, no Excel-DNA dependency in the core so it tests cross-platform).
  - `StateService` — process-wide registry of immutable `WorkbookState` records keyed by workbook id, single coarse lock for mutation, `StateChanged` event for ribbon-invalidation hooks. Typed helpers for every ribbon edit (SetEnabled / SetSelectedScript / AddAction (upsert) / DeleteAction / …).
  - `WorkbookState` + `RibbonAction` records — immutable, with-clone-based mutation, `SelectedAction` convenience accessor.
  - `IWorkbookContext` — abstraction over "which workbook is active", `NullWorkbookContext` default keeps the ribbon sane before the add-in's `AutoOpen` wires the real implementation.
  - `ScriptDirectoryWatcher` — wraps `FileSystemWatcher` against a `userScripts/` dir, normalises the `.py` list (deduped, sorted, name-only), pushes an initial snapshot synchronously so consumers don't sit at empty until the first edit lands.
  - `PyExcelServices` static service locator so Excel-DNA's parameterless-ctor `PyExcelRibbon` can pull dependencies without DI gymnastics.
  - `PyExcelRibbon`: every formerly-hardcoded getter (`RibbonEnabled`, `GetScriptCount`, `GetScriptLabel`, `GetScriptText`, `GetPyInput`, `GetPyOutput`, `GetActionCount`, `GetActionLabel`, `GetActionText`) reads from `PyExcelServices.State.Get(activeKey)`. `OnPyInputChange` / `OnPyOutputChange` / `OnScriptChange` / `OnActionChange` write back. `OnDeleteAction` now actually deletes. `OnAddAction` / `OnEditAction` still stubbed pending the EditActionForm (Phase 8); their state plumbing is in place. `RibbonOnLoad` subscribes to `StateChanged` and queues `IRibbonUI.Invalidate` via `ExcelAsyncUtil.QueueAsMacro` so changes originating from a worker thread (`FileSystemWatcher`, Excel COM events, ...) refresh the ribbon safely.
  - Tests: 17 `StateServiceTests` covering empty defaults, mutator-preserves-fields, multi-workbook isolation, Forget semantics, Action upsert / delete-and-clear-selection, argument validation. 7 `ScriptDirectoryWatcherTests` against a tmp directory exercising initial snapshot, add/delete/rename events, manual `Refresh`, non-`.py` filter, argument validation, idempotent `Dispose`.
  - The CustomXMLPart persistence + `AppEventSink` (COM event sink) + the routing of Excel-DNA UDF cancellation through `KernelClient.Cancel` are the COM-coupled tail; tracked in `docs/phase3-and-4-completion.md` for the next session.
- **2026-05-24 — Phase 4 step 5: cooperative CANCEL.** The supervisor now dispatches `RUN_REQUEST` on a worker thread and pumps inbound frames in the main loop, so `CANCEL` (and `PING`) arriving during a run actually land. The cancel flag is exposed to user code as `pyexcel.kernel.is_cancelled()`; long-running scripts can poll it between work units and return early. Either way (cooperative abort or natural completion), the kernel surfaces `ERROR / Cancelled` so the host can tell the run was interrupted. Two new module-level functions in `worker.py` (`_begin_job` / `_end_job` / `is_cancelled`), one new method on `FrameTransport` (`has_data(timeout_s)`) backed by `select` on POSIX and a `PeekNamedPipe` ctypes call on Windows, plus the `_run_with_cancellation` helper in `supervisor.py`. Three end-to-end pytest tests cover: CANCEL during a long run → `ERROR / Cancelled`, unsolicited CANCEL → polite `ERROR` + loop stays alive, PING during a run → `PONG` arrives. 140/140 kernel tests green. The Excel-DNA UDF doesn't route Excel-DNA's task cancellation to `KernelClient.Cancel` yet — that bridge is a follow-up.
- **2026-05-24 — Phase 4 step 4: error visibility via LogDisplay.** `PyRunFunction` now writes kernel-error details (`Code`, `PythonType`, `Message`, full `PythonTraceback`) to `ExcelDna.Logging.LogDisplay` in addition to `Trace`. LogDisplay is Excel-DNA's built-in pop-up log window — surfaces in Excel itself rather than only in DebugView. Cell still gets `#VALUE!` so `ISERROR()` keeps working. No code change required of users; just a packaging upgrade to a more useful error path.
- **2026-05-24 — Phase 4 step 3: SAFE-1 async UDF.** Replaced the sync UDF body with `ExcelAsyncUtil.Run(...)` so the calc thread no longer blocks for the duration of a kernel run. First call returns `#N/A` immediately; Excel-DNA refreshes the cell when the background task completes. Identical inputs short-circuit to the cached result instead of re-spawning. The blocking work now runs in `RunSynchronously` on Excel-DNA's worker thread; same error contract as before (KernelException → `#VALUE!` + `Trace.WriteLine` of the Python traceback). Worker-side cancellation (kernel acting on CANCEL frames) is still pending — Excel-DNA cancels its background task on formula change, but the kernel completes the run regardless.
- **2026-05-24 — Phase 4 step 2: `=PY.RUN` UDF.** Added `PyRunFunction.cs` (net48-only, `#if NETFRAMEWORK` so the netstandard2.0 / Linux CI slice ignores it). ExcelDna.Integration 1.8.0 conditional package reference for the net48 target. The UDF is a small translation layer over `PyRun.Execute` — Excel sentinel arguments map to `null`, the `PyRun.EmptyResult` sentinel maps to `ExcelEmpty.Value`, `KernelException` is rendered as `#VALUE!` in the cell with the Python traceback logged via `Trace.WriteLine`. Synchronous in this slice; async/progress/cancel (SAFE-1) and the ribbon button are separate Phase 4 items. No automated test coverage on this file alone — exercising `[ExcelFunction]` needs an actual Excel instance — but the dispatch core it delegates to is covered by `PyRunTests`' 13 end-to-end cases.
- **2026-05-24 — Phase 4 step 1: marshalling + dispatch core.** New `PyExcel.Excel` assembly (multi-target net48 + netstandard2.0, Apache.Arrow 18.0.0). Three pieces:
  - `ArrowMarshal` — C# half of the kernel data plane. `EncodeTable` / `EncodeVector` / `EncodeScalar` / `PeekShape` / `Decode`. Schema metadata (`pyexcel-shape`, `pyexcel-orientation`) matches `arrow_io.py` byte-for-byte. Per-column type inference (double/bool/string with string fallback for mixed), nulls preserved. 23 unit tests.
  - `PythonResolver` + `KernelHost` — discovery (env var → workbook venv → PATH) plus a process-wide `Lazy<KernelClient>` whose first access boots the kernel and whose `Dispose` is the add-in unload hook. Phase 3 will move ownership to per-workbook state.
  - `PyRun.Execute(script, input, kwargs, client, …)` — shared marshal-and-dispatch core for both the planned UDF and the ribbon button. 13 e2e tests through a real Python kernel are the cross-language conformance check for ArrowMarshal ↔ arrow_io.py.
  - Repository housekeeping: previously fragmented `claude/*` branches consolidated into the single `NET_Migration` working branch.
- **2026-05-23 — Phase 2 complete.** Shipped the rest of the kernel data plane in one session, both CI lanes green on `ae0b4f0`:
  - `arrow_io.py` — shape-preserving Arrow IPC for DataFrame / Series / list / tuple / 1-D-or-2-D numpy / scalar, with `pyexcel-shape` and `pyexcel-orientation` schema metadata so the host can reconstruct cell geometry. 39 pytest tests.
  - `worker.py` — pure `run_job(meta, payloads) -> JobOutcome`; loads the user script (mtime-cached), decodes Arrow payloads, calls the target function, replies with `RUN_RESULT` or a typed `ERROR` (9 stable codes: `BadRequest` / `ModuleNotFound` / `ModuleLoadError` / `ModuleExecError` / `FunctionNotFound` / `FunctionNotCallable` / `BadInput` / `BadReturnType` / `Exception`). 23 unit tests + 2 e2e supervisor tests.
  - `PyExcel.Kernel.Client` — new assembly. Typed `RunRequest` / `RunResult` / `KernelException`, `ProgressReceived` / `LogReceived` events, sync `Run` + async `RunAsync`, fire-and-forget `Cancel`. Required a dual-lock refactor of `KernelSupervisor` (`ExchangeSemaphore` + separate read/write locks) so Cancel can fire while a Run is parked in a read. 15 C# integration tests against a real Python subprocess.
  - Windows named-pipe transport — Python `_winapi` client against `\\.\pipe\<name>` with retry on `ERROR_PIPE_BUSY` / `ERROR_FILE_NOT_FOUND`. C# side now sets a DACL granting only the current-user SID at pipe creation (net48 only — netstandard2.0 is Linux-only in CI, so the DACL block is `#if NETFRAMEWORK`).
  - CI: Windows lane now installs Python 3.12 + pyarrow/pandas/numpy and runs the C# integration tests against the Win32 transport. Both lanes green; Phase 2 exit criteria fully met.
- **2026-05-23 — Phase 2 step 3: KernelSupervisor + python entry point.** Added `KernelSupervisor.cs` (C# owns the named-pipe server, spawns `python -m pyexcel.kernel --pipe <name>` via argv, runs HELLO handshake with protocol-version check, exposes `Ping`/`Shutdown`, and force-kills on dispose so no orphaned python.exe). Added `embedded/pyexcel/kernel/{transport.py, supervisor.py, __main__.py}` — POSIX/AF_UNIX client connecting to `/tmp/CoreFxPipe_<name>` (matches .NET's pipe-on-Linux path), supervisor event loop handling HELLO/PING/PONG/SHUTDOWN, ERROR reply for not-yet-supported frames. Tests: 3 C# integration tests (round-trip + 10×PING + dispose-without-shutdown) and 3 pytest tests (handshake roundtrip, protocol-mismatch rejection, ERROR-after-unsupported-frame keeps loop alive). CI workflow reordered so Python is set up before `dotnet test` (integration tests spawn python). Windows transport stub raises `NotImplementedError` — landing alongside the Windows kernel CI slice in a later step.
- **2026-05-23 — Phase 2 step 2: FrameTransport (stream + named-pipe).** Added `FrameTransport.cs` wrapping any `Stream` (MemoryStream for tests, `NamedPipeClientStream` for production) with synchronous `ReadFrame`/`WriteFrame` and a `ConnectNamedPipe` static factory. Added `FrameTransportTests.cs` covering MemoryStream roundtrip, dispose semantics, oversize rejection, and a real Windows-named-pipe/Linux-Unix-domain-socket roundtrip pairing client to in-process server.
- **2026-05-23 — Phase 2 step 1: framing.** Added `PyExcel.Bridge` (multi-target `net48`/`netstandard2.0`) with `Framing.cs` + a stdlib `CanonicalJson` encoder/decoder, mirroring `framing.py` byte-for-byte. Added `PyExcel.Bridge.Tests` (xUnit, net8.0) with the Python test-suite ported 1:1, plus cross-language golden hex vectors (`test_cross_language_vectors.py` ↔ `CrossLanguageVectorsTests.cs`) that pin the on-wire format. CI now builds the netstandard slice on Linux and runs `dotnet test`.
- **2026-05-22 — Phase 0 complete.** Removed the personal-data `__main__` block and trailing dead code from `xmlParsing.py`; replaced the bloated root `requirements.txt` with the minimal v2-kernel set; added `.github/workflows/ci.yml`; confirmed the v1-frozen policy below.
- **2026-05-22 — Planning.** Repo audited; `ARCHITECTURE.md` and this roadmap written; version/phase mismatches reconciled; the four architecture decisions resolved.

**v1 maintenance policy:** v1 (`PyExcel.xlam`) is frozen while v2 is built — security and critical-data-loss fixes only, no new features. It remains the shipping product until Phase 9 cutover.

## Definition of production-grade (the bar every phase must clear)

A phase delivery is complete only when **all** of these hold for the code it adds:

- [ ] **Tested.** Unit tests land *with* the code; cross-language contracts (framing, Arrow) have conformance tests. No module merged without tests.
- [ ] **No dead code.** No commented-out blocks, no unused helpers, no "OLD" files, no debug-print artifacts.
- [ ] **Errors are real.** No silent swallowing (`On Error Resume Next` / empty `catch`). Every failure is logged via `ILog`; user-facing failures are surfaced with an actionable message.
- [ ] **No resource leaks.** Processes, COM objects, file handles, pipes, timers are deterministically released.
- [ ] **Builds clean.** `TreatWarningsAsErrors` is on — zero warnings. `dotnet build` + `pytest` green in CI.
- [ ] **Documented.** [`ARCHITECTURE.md`](ARCHITECTURE.md) and this file are updated as part of the phase's definition of done.
- [ ] **No hardcoded paths or personal data.** Nothing machine- or author-specific.

## Canonical phase model

This is the **single source of truth** for phase numbering. `docs/v2-build.md`
and the `// PHASE n` comments in the C# defer to it.

| Phase | Title | Headline deliverable |
| --- | --- | --- |
| 1 | Skeleton | `.xll` loads, ribbon tab renders ✅ |
| 2 | Bridge & kernel core | C# ↔ Python talk over the pipe |
| 3 | State & events | Ribbon reflects per-workbook state |
| 4 | Excel marshalling & first run | One real script runs end to end |
| 5 | Data services & shell | Import / Export / Paste work |
| 6 | Charts | Plotly/Matplotlib results render as Excel charts |
| 7 | Setup | Fresh-machine provisioning works |
| 8 | Forms & UI polish | Every dialog is a working WinForm |
| 9 | Cutover & v1 retirement | v2 ships; VBA deleted |

---

## Phase 0 — Reconciliation & cleanup (do before Phase 2)

Make the repo internally consistent and remove what must never reach v2.

- [x] Remove the `panadas==0.2` typosquat from `requirements.txt`.
- [x] Re-encode `src/embedded/requirements.txt` from UTF-16 to UTF-8.
- [x] Reconcile version strings (`README.md`) and phase numbers (`docs/v2-build.md`, C# `// PHASE` comments).
- [x] Add `ARCHITECTURE.md` and this `ROADMAP.md` as the canonical reference docs.
- [x] **Strip personal data** — removed the `__main__` test block and trailing commented-out dead code from `src/embedded/xmlParsing.py` (1941 → 762 lines).
- [x] **Slim dependencies** — root `requirements.txt` is now the minimal v2-kernel set (`pandas numpy pyarrow plotly matplotlib`); the author-specific `git+https` deps are dropped. `src/embedded/requirements.txt` stays as v1's frozen install set until Phase 9.
- [x] **v1 maintenance policy** — confirmed frozen: security / critical-data-loss fixes only, no new features (see *Status & progress log*).
- [x] **CI workflow** — `.github/workflows/ci.yml`: Linux builds the `netstandard2.0` slice + runs `pytest`; Windows builds the full solution.

---

## Phase 1 — Skeleton ✅ (complete)

Delivered: `PyExcel.Common` (logging), `PyExcel.Ribbon` skeleton, `PyExcel.Addin`
skeleton, `embedded/pyexcel/kernel/framing.py` + tests. The `.xll` loads, the
Python tab renders, every button is a logged stub.

---

## Phase 2 — Bridge & kernel core

**Objective.** C# and the Python kernel exchange frames over a named pipe; the
kernel's lifetime is owned and deterministic.

**Deliverables.** `PyExcel.Bridge`, `PyExcel.Kernel.Client`, and the Python
`pyexcel.kernel` package (`supervisor.py`, `worker.py`, `arrow_io.py`, `__main__.py`).

- [x] `PyExcel.Bridge/Framing.cs` — mirror `framing.py` byte-for-byte (same frame layout, bounds, determinism).
- [x] Cross-language conformance tests: frames encoded in C# decode in Python and vice versa (golden hex vectors pin the wire format on both sides).
- [x] Bounded/malformed-frame handling on the C# side (mirror the `framing.py` test suite).
- [x] Named-pipe transport — POSIX side (C# `NamedPipeServerStream` ↔ Python `socket(AF_UNIX)` against `/tmp/CoreFxPipe_<name>`) and Windows side (C# named pipe with a DACL granting only the current-user SID ↔ Python `_winapi` client against `\\.\pipe\<name>`). Wrong-user processes get `ERROR_ACCESS_DENIED` at connect time before any frame bytes cross the boundary.
- [x] `KernelSupervisor` — spawns `python -m pyexcel.kernel` via argv (no shell), `HELLO` handshake with protocol-version check, `Ping`/`Shutdown` API, deterministic kill on `Dispose`. PING/PONG health-check loop on a background timer is a separate item.
- [x] `PyExcel.Kernel.Client` — typed API: `RunRequest`/`RunResult`/`KernelException`, `ProgressReceived`/`LogReceived` events, sync `Run` + async `RunAsync`, and a fire-and-forget `Cancel` that uses only the write lock so it can fire while a `Run` is parked in a read. Requires a dual-lock refactor of `KernelSupervisor` (exchange semaphore + separate read/write locks) — done in the same change.
- [x] Python `supervisor.py` — connects to the pipe, runs the HELLO/PING/PONG/SHUTDOWN loop. Worker dispatch is a follow-up.
- [x] Python `worker.py` — run one job: receive `RUN_REQUEST` → load the user script (mtime-cached) → decode Arrow payloads → call the target function → reply with `RUN_RESULT` or a typed `ERROR` (`BadRequest` / `ModuleNotFound` / `ModuleExecError` / `FunctionNotFound` / `FunctionNotCallable` / `BadInput` / `BadReturnType` / `Exception`). Pure function; supervisor wires it into the dispatch loop.
- [x] Python `arrow_io.py` — DataFrame / list / scalar ↔ Arrow IPC stream. Shape (`table`/`vector`/`scalar`) and vector orientation are carried as Arrow schema metadata so the host can spill back into the right cell geometry.

**Exit criteria.** C# spawns the kernel, completes a `HELLO`/`PING`/`PONG`
round-trip, and kills it cleanly on shutdown. Framing conformance tests pass in
both languages. No orphaned `python.exe` after Excel closes.

---

## Phase 3 — State & events

**Objective.** The ribbon accurately reflects per-workbook state; switching
sheets/workbooks never shows stale values.

**Deliverables.** `PyExcel.State`, an `AppEventSink`.

- [x] `StateService` — new `PyExcel.State` assembly. Immutable `WorkbookState` records keyed by workbook id; thread-safe `Get` / `Update` / typed setters (`SetEnabled`, `SetSelectedScript`, `AddAction`, `DeleteAction`, …) plus a `StateChanged` event. `IWorkbookContext` abstraction over `Application.ActiveWorkbook`, with a `NullWorkbookContext` default so the ribbon renders sanely before `AutoOpen` wires the real one. Pure-logic; tested cross-platform.
- [x] Persist per-workbook state as a `CustomXMLPart` on the workbook. `WorkbookStatePersister` (net48, COM) writes on `WorkbookBeforeSave` / reads on `WorkbookOpen`, delegating the XML to the tested `WorkbookStateCodec`. Code-complete; runtime behaviour pending the Windows smoke test (`docs/phase3-and-4-completion.md` §3).
- [x] `AppEventSink` — `WorkbookOpen` / `Activate` / `BeforeSave` / `BeforeClose` / `SheetActivate` → update state + invalidate ribbon. Typed Application events via the embedded `ExcelDna.Interop` PIA (CI-restorable, no runtime dependency). Code-complete; runtime behaviour pending the Windows smoke test (`docs/phase3-and-4-completion.md` §2).
- [x] Wire `RibbonEnabled` / all getters to `StateService`. Every former-hardcoded callback in `PyExcelRibbon` now reads from `PyExcelServices.State.Get(activeKey)`. `OnPyInputChange` / `OnPyOutputChange` / `OnScriptChange` / `OnActionChange` write back through `StateService`. Ribbon redraw queued via `ExcelAsyncUtil.QueueAsMacro` so `StateChanged` from a worker thread is safe.
- [x] `FileSystemWatcher` on `userScripts/` — new `ScriptDirectoryWatcher` watches a directory for `.py` files and pushes a sorted, deduped list to a callback (typically wired to `StateService.SetAvailableScripts`). Initial snapshot fires inside the constructor so consumers start with the current contents. Tested cross-platform with `tmp` directories.
- [~] Wire `OnAddAction` / `OnEditAction` / `OnDeleteAction`. **`OnDeleteAction` is wired** to `StateService.DeleteAction` against the selected action. Add/Edit need the EditActionForm (Phase 8); the state plumbing (`StateService.AddAction` accepts upsert) is ready to receive form output.

**Exit criteria.** Enabling a workbook lights the ribbon; switching sheet or
workbook updates every ribbon field correctly; state survives a save/reopen.

---

## Phase 4 — Excel marshalling & first run (the thin slice)

> **Handoff brief:** [`docs/phase4-handoff.md`](docs/phase4-handoff.md) —
> what Phase 2 shipped (public APIs, wire contract, known gaps) and what
> Phase 4 still has to build (C#-side Arrow encoder, `=PY.RUN` UDF,
> short-term kernel lifecycle). Read that first.

**Objective.** Prove one full run works end to end — this de-risks the pipe,
Arrow, COM interop, and the threading model in a single slice before going wide.

**Deliverables.** `PyExcel.Excel`; a wired `OnRunPython`.

- [x] Range → Arrow: read `object?[,]` / `object?[]` / scalar into Arrow IPC streams with shape metadata (`pyexcel-shape`, `pyexcel-orientation`) matching `arrow_io.py`. Per-column type inference; mixed columns fall back to string. Date/Excel-error handling still ahead.
- [x] Arrow → Range: decode Arrow IPC back to `object?[,]` / `object?[]` / scalar. Defaults to table for buffers without PyExcel metadata (interoperable with external Arrow writers).
- [x] `PyRun.Execute` — shared marshal-and-dispatch core for both the `=PY.RUN` UDF and the ribbon `OnRunPython` button. Resolves the script path (relative → workbook dir), encodes input as Arrow, calls `KernelClient.Run`, decodes the response (honours `pyexcel-orientation` to spill row vs column), and surfaces a `None` return as a sentinel the wrapper translates to `ExcelEmpty`. No Excel-DNA dependency — runs in netstandard2.0, fully unit-testable.
- [x] `KernelHost` — process-wide `Lazy<KernelClient>` lifecycle wrapper for Phase 4. First `Client` access boots the kernel; idempotent `Dispose` for the add-in unload hook. Phase 3 will replace this with per-workbook ownership in `StateService`.
- [x] `PythonResolver` — three-tier discovery: `PYEXCEL_PYTHON` env var → `<workbook>/.pyexcel-venv/{Scripts,bin}/python` → PATH fallback. Plus `ResolveEmbeddedPath()` walking up from `AppContext.BaseDirectory` to find `embedded/pyexcel/kernel/__main__.py`.
- [ ] Parse the Input/Output ribbon fields, including the `{name}=Range` syntax.
- [x] `=PY.RUN(script, input, [function])` Excel-DNA UDF (net48-only) — wrapper around `PyRun.Execute` that translates Excel-DNA sentinel args (`ExcelMissing`/`ExcelEmpty`/`ExcelError`) to plain .NET shapes and threads errors back as `#VALUE!` with the full Python traceback logged to `Trace`.
- [x] **SAFE-1** for the UDF: dispatched through `ExcelAsyncUtil.Observe`, so the calc thread never blocks. First call returns `#N/A` immediately; the cell auto-refreshes when the kernel returns. (The earlier slice used `ExcelAsyncUtil.Run`; the switch to `Observe` is what enables the UDF cancel bridge below.)
- [x] Kernel-side cooperative CANCEL — supervisor dispatches `RUN_REQUEST` on a worker thread and pumps inbound frames during the run, so `CANCEL` sets a flag the user's `transform()` can read via `pyexcel.kernel.is_cancelled()`. `PING` continues to answer for liveness. The kernel replies `ERROR / Cancelled` whether the user code noticed and aborted or completed naturally. **Now also wired through the UDF:** the `=PY.RUN` observable cancels its `CancellationTokenSource` on dispose; the token registration inside `KernelClient.RunAsync` pushes the `CANCEL` frame, so Excel-DNA's cancel-on-formula-change actually aborts the run. See `docs/phase3-and-4-completion.md` §4.
- [x] `OnRunPython` ribbon button handler (same dispatch core as the UDF, called from the button event), non-blocking SAFE-1 pattern. Shipped + verified on real Excel 2026-05-29 (see status log / `docs/phase3-and-4-completion.md` §7). Progress UI is the item below.
- [x] Kernel-side progress: `pyexcel.kernel.report_progress(percent, message)` emits `PROGRESS` frames the supervisor streams to the host (`KernelClient.ProgressReceived` already wired). 6 tests; 146/146 kernel tests green.
- [ ] Non-blocking progress UI (WinForms) rendering those `PROGRESS` frames, with a working **Cancel**. Windows-only; Phase 8.
- [x] Archive a run (inputs, outputs, log); retention cap. `PyExcel.State.RunArchive` writes each run to `%LOCALAPPDATA%\PyExcel\runs\{yyyyMMddTHHmmssfff}_{rand}\` — Arrow IPC bytes for every input (`input_N.arrow`), the output (`output.arrow`, present iff the run produced a payload), a human-readable manifest (`manifest.txt`, line-per-field), and the formatted error block (`error.txt`, present on Error / Cancelled). Retention defaults to 20 most-recent runs; older directories are evicted FIFO after each write. Threaded through both surfaces via a new optional `RunArchiveContext` parameter on `PyRun.Execute*` — `PyRunFunction` tags its runs `"PY.RUN"`, `RangeRunner` tags them `"Run Python button"`. Best-effort: an I/O fault inside the archive can't mask the user-facing result or the in-flight exception.
- [~] Surface kernel errors — `KernelException` and host-side faults now flow into a per-workbook `ErrorService` (with a global fallback slot for failures that arrive before any workbook is bound). The ribbon adds an **Errors** group with **Show Last Error** (opens Excel-DNA's `LogDisplay` with the formatted traceback) and **Copy Last Error** (puts the same block on the clipboard for bug reports). Both buttons enable only when there's something to show — `RibbonHasError` + `ErrorService.ErrorChanged` repaint hook. Cell still shows `#VALUE!` so `ISERROR()` keeps working. Still outstanding: a per-script error pane and a hover tooltip on the failing cell — both are Phase 8 WinForms / UI work.
- [x] `ExcelFormula` round-trip (A1 mode) — marshalling layer only. Python `pyexcel.kernel.Formula("=…")` ↔ C# `PyExcel.Excel.Formula("=…")`. Wire format: a string Arrow column with field-level metadata `pyexcel-cell-type=formula`. Scalar + vector are wired; DataFrame-with-formula-columns is skipped (encode side flattens to plain strings — documented gap). The RangeRunner write hook (`range.Formula = …` vs `range.Value2 = …`) is a separate ribbon-integration follow-up.
- [x] Tests: range ↔ Arrow round-trip incl. dates, nulls, formulas. **Note:** writing the date tests surfaced and fixed a real decode bug — `ArrowMarshal.ReadCell` was falling through to `_ => array.GetType().Name` for `Date32Array` / `Date64Array` / `TimestampArray`, so a Python script returning a `datetime.datetime` rendered the literal string `"TimestampArray"` in Excel. The decoder now lands on a naive `DateTime` (matching pyarrow's no-timezone default for `datetime` / `date`).

**Exit criteria.** A user `transform()` taking one input table and returning one
table runs from a ribbon click and writes results; a failing script shows a
clear, actionable error.

---

## Phase 5 — Data services & shell

**Objective.** Import, Export, and Paste work correctly for the documented formats.

**Deliverables.** Import/Export/Paste services in `PyExcel.Excel`; `PyExcel.Common.Shell`.

- [ ] CSV/TSV import via a real **RFC-4180** parser — embedded newlines/commas/quotes, UTF-8 + BOM detection, correct delimiter for `.tsv`.
- [ ] Excel-format import (XLSX/XLSM/XLSB/ODS) via COM; sheet picker that lists the **source** workbook's sheets.
- [ ] Export ranges to CSV/Excel with correct quoting and explicit encoding.
- [ ] Paste a saved artifact to a range, with overwrite confirmation.
- [ ] `PyExcel.Common.Shell` — open Explorer, open the user's editor, open the readme; wire `OnOpenExplorer` / `OnReadMe` / `OnEditPython`.
- [ ] Wire the Import / Export / Paste ribbon groups end to end.
- [ ] Tests for CSV edge cases.

**Exit criteria.** Import/export/paste handle the documented formats with correct
encoding and quoting; no data corruption on round-trip.

---

## Phase 6 — Charts

**Objective.** A `transform()` that returns a figure produces a native Excel chart.

**Deliverables.** `PyExcel.ChartBuilder`; chart support in the kernel.

- [ ] Kernel: Plotly figure → a JSON **chart spec** (port the `PlotlyToExcelXMLConverter` traversal; emit JSON, not XML — carried in the `RUN_RESULT` frame).
- [ ] Kernel: Matplotlib figure → image artifact (SVG, PNG fallback).
- [ ] `PyExcel.ChartBuilder`: JSON chart spec → native Excel chart (port the live chart-COM logic from `chartBuilder.bas`).
- [ ] Guard against orphan charts (clean up on failure) and missing spec attributes.
- [ ] Support the documented chart types; explicit, surfaced handling of unsupported types.
- [ ] Tests for the chart-spec contract.

**Exit criteria.** A Plotly figure renders as a native Excel chart; a Matplotlib
figure embeds as an image; a malformed spec fails cleanly with a message.

---

## Phase 7 — Setup

**Objective.** A fresh machine can be provisioned reliably, with diagnosable failures.

**Deliverables.** `PyExcel.Setup`.

- [ ] First-run wizard — project folder, convert host to `.xlsm`, create the project tree.
- [ ] venv creation against system Python; detect a missing interpreter or the Windows Store stub and say so.
- [ ] Extract the kernel package from **.NET embedded resources** (no base64 sheet, no chunk-assembly).
- [ ] `pip install` the canonical requirements; capture stdout+stderr to a real log; surface failures.
- [ ] Dependency verification with a clear pass/fail (no silent 80% threshold).
- [ ] Path resolution that handles UNC (`\\server\share`) and localized / non-default SharePoint libraries.
- [ ] Retire `Update.bas` — updating = shipping a new `.xll` (Excel-DNA reloads it).

**Exit criteria.** Clean setup on a fresh Windows machine; every failure mode
produces a specific, actionable message; UNC and SharePoint project paths work.

---

## Phase 8 — Forms & UI polish

**Objective.** Every ribbon action that needs a dialog has a working one.

**Deliverables.** `PyExcel.Forms` (WinForms).

- [ ] Rewrite the 9 dialogs as WinForms: range picker, edit action/import/export/paste, export wizard, orientation, sheet picker, progress.
- [ ] Wire **every** control — fix the dead `frmExportWizard` row buttons.
- [ ] Replace the off-screen `(-20000,-20000)` form-hide hack with proper modal/owner handling.
- [ ] Input validation in every dialog; clear inline error messaging.
- [ ] Ship the ribbon logo PNG as an embedded resource (`LoadImage`).

**Exit criteria.** No unwired controls; no dialog can be lost off-screen; invalid
input is caught in the dialog, not downstream.

---

## Phase 9 — Cutover & v1 retirement

**Objective.** v2 becomes the product; v1 is removed.

- [ ] Feature-parity checklist vs v1 — every documented capability verified in v2.
- [ ] v1 → v2 per-workbook state migration (or a documented re-enable path).
- [ ] End-to-end QA on Excel 2016 / 2019 / 365 (x64).
- [ ] Delete `src/module/`, `src/embedded/`, `src/Ribbon/`, `PyExcel.xlam`.
- [ ] Rewrite `README.md` to describe v2 as the product.
- [ ] Tag the release; update `Directory.Build.props` version off `-alpha`.

**Exit criteria.** v2 ships; no VBA remains in the repository.

---

## Bug disposition

The full audit, classified for the migration.

### Designed out — the v2 architecture removes these; do not port

Locale-dependent number text · `CLng` integer overflow · Excel-epoch/date
heuristics · column/scalar type-inference mistakes · `meta.xml` polling hangs,
heartbeat false-stalls, run-id mismatch · orphaned `python.exe` · `cmd /c`
quoting/injection · multi-chunk base64 extraction corruption · UTF-16
`requirements.txt` · last-sheet-row format clobber.

### Fix in transit — must be done deliberately during the port

- [ ] CSV parsing — adopt a real RFC-4180 library (Phase 5).
- [ ] Silent failure — `ILog` + surfaced errors everywhere (every phase; enforced by the production-grade bar).
- [ ] Stale sheet state — single source of truth in `PyExcel.State` (Phase 3).
- [ ] Subprocess hardening — `KernelSupervisor` owns lifetime (Phase 2).
- [ ] `frmExportWizard` dead buttons — wire them in the WinForms rewrite (Phase 8).
- [ ] chartBuilder orphan charts / null-attribute crashes — guard in `PyExcel.ChartBuilder` (Phase 6).
- [ ] Destructive paste — add overwrite confirmation (Phase 5).
- [ ] Setup diagnostics — surface venv/pip output (Phase 7).

### Fix now — handled in Phase 0

Typosquat removal ✅ · UTF-16 requirements ✅ · dependency slimming ✅ · personal-data removal ✅

## Cut list — delete, do not migrate (~3,000+ lines)

- [ ] `xmlParsing.bas` — entire file (Arrow replaces it).
- [ ] `chartBuilder.bas` lines ~2–1010 — dead alternate engine.
- [ ] `tools.py` — ~1,200 commented-out legacy lines.
- [x] `xmlParsing.py` — ~485 commented lines + the ~570-line `__main__` block. *(done — Phase 0)*
- [ ] `frmEditActionOLD.frm/.frx`; commented blocks in the three `frmEdit*` forms; the `CAppEvents.cls` duplicate.
- [ ] `Import.bas ReadExcel_ADO` (~100 dead lines); unused helpers across modules.
- [ ] Most of `Update.bas` (SmartClean / manifest / version-name machinery).
- [ ] `PyExcel.xlam` and all of `src/module/` — at Phase 9.

## Cross-cutting — CI & testing

- [ ] Add `.github/workflows/` — on every push: build the `PyExcel.Common` netstandard2.0 slice + run `pytest tests/` (Linux); build the full solution on a Windows runner.
- [ ] Gate merges on green CI.
- [ ] Every phase lands with tests; track coverage of `PyExcel.Excel`, `PyExcel.Bridge`, and the kernel.
- [ ] Code review on every PR against the *Definition of production-grade*.

## Decisions (resolved)

These were open; now settled. Recorded here as the source of truth.

1. **Chart transport — JSON chart spec.** The kernel converts Plotly figures to a JSON chart spec carried in the `RUN_RESULT` frame; `PyExcel.ChartBuilder` builds the native Excel chart from it. No XML layer. *(Phase 6)*
2. **List/scalar transport — everything as Arrow.** Lists become 1-column Arrow batches and scalars 1×1 batches — one uniform marshalling path with full type fidelity, including timestamps. *(Phase 2 `arrow_io.py`, Phase 4)*
3. **State storage — `CustomXMLPart`.** Per-workbook state is stored as a workbook-attached `CustomXMLPart`: invisible, no length limits, no Name Manager clutter. Phase 9 needs a converter that reads v1's defined Names and writes the new part. *(Phase 3, Phase 9)*
4. **Forms UI — WinForms.** The 9 dialogs are rebuilt as WinForms — lowest hosting friction with Excel-DNA, near 1:1 with the existing layouts. *(Phase 8)*
