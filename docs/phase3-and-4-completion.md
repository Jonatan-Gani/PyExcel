# Phase 3 + Phase 4 completion brief

> **Audience.** A fresh session picking up where this branch
> (`NET_Migration`) stopped. Skim this once; [`ROADMAP.md`](../ROADMAP.md) and
> the existing tests are the source of truth. The companion brief
> [`docs/phase4-handoff.md`](phase4-handoff.md) covers the Phase 2 → 4
> transition and the kernel data plane that's already shipped.

---

## TL;DR

The cross-platform foundation for Phases 3 and 4 is **complete and CI-green**:

- The Python kernel runs jobs end-to-end with cooperative `CANCEL` support.
- `=PY.RUN` is wired (Excel-DNA UDF), async via `ExcelAsyncUtil.Run`,
  errors surface in Excel-DNA's `LogDisplay`.
- `StateService` owns per-workbook ribbon state; the ribbon reads it.
- `ScriptDirectoryWatcher` keeps the Script dropdown live against
  `userScripts/`.

What's **left** is the layer that touches **Excel COM** — and is therefore
impossible to exercise from Linux CI. Everything in this list needs a
real Excel install on Windows to verify by hand. The components are small;
the friction is in the test loop, not the code.

---

## What's left to fully close Phase 3 + Phase 4

### 1. `ExcelWorkbookContext` — concrete `IWorkbookContext` ✅

Landed in `src/PyExcel.Addin/ExcelWorkbookContext.cs`. Uses
`ExcelDnaUtil.Application` over `dynamic` (no Office PIA reference).
Saved workbooks → `Workbook.FullName`; unsaved workbooks →
`"unsaved:{SessionGuid}:{Workbook.Name}"` (the session GUID is
allocated once per add-in load so two unsaved books named `Book1` /
`Book2` don't collide). Any COM exception during the lookup
(transient state between workbook events, shutdown) is swallowed and
yields `null` — the ribbon's getters all tolerate that as
"no workbook".

Wired into `PyExcelServices.WorkbookContext` from `AddIn.AutoOpen`.
Can't unit-test on Linux (no Excel COM); manual smoke test on
Windows: open a fresh workbook → save it → close → reopen, and
confirm the ribbon's getEnabled / getSelectedScript reflect the
right per-workbook state at each step.

### 2. `AppEventSink` — the Excel.Application event subscriber ✅ (code-complete — needs smoke test)

Landed in `src/PyExcel.Addin/AppEventSink.cs` (net48). Subscribes to
`Application.WorkbookOpen`, `WorkbookActivate`, `WorkbookBeforeSave`,
`WorkbookBeforeClose`, `SheetActivate`. Every handler body runs inside a
`Guard(name, action)` wrapper so a fault is logged and swallowed, never
propagated back into Excel's event pump. Constructed in `AddIn.AutoOpen`
(inside its own try/catch so a sink failure degrades gracefully — the
ribbon and `=PY.RUN` keep working), disposed in `AddIn.AutoClose`
(unsubscribes every handler). Each handler:

- `WorkbookOpen(wb)` → restore state from CustomXMLPart (#3),
  `ScriptDirectoryWatcher` for that workbook's `userScripts/`,
  `IRibbonUI.Invalidate`.
- `WorkbookActivate(wb)` → `IRibbonUI.Invalidate` (the workbook key
  changed; every getter has a new state to render).
- `WorkbookBeforeClose(wb)` → flush state to CustomXMLPart (#3),
  `StateService.Forget(key)`, dispose the per-workbook watcher.
- `SheetActivate(sh)` → `StateService.SetCurrentSheet(key, sh.Name)`.

Implementation reference: Excel-DNA's
[`AddInBase`](https://github.com/Excel-DNA/ExcelDna/wiki/Excel-DNA-AddIns-Application-Object)
or `Application` COM via `ExcelDnaUtil.Application`. Lifetime is the
add-in lifetime (`AutoOpen` → `AutoClose`).

> **✅ Design decision — resolved: typed PIA via `ExcelDna.Interop`,
> embedded.** Subscribing to `Application.WorkbookOpen` &c. needs a typed
> delegate (`app.WorkbookOpen += handler`) — you can't `+=` over a
> late-bound `dynamic`, which is why every *other* COM-bound piece
> (`ExcelWorkbookContext`, `RangeRunner`) stays late-bound but this one
> can't. The original worry was that a typed
> `Microsoft.Office.Interop.Excel` reference breaks the Windows CI build,
> since `windows-latest` has no Excel and CI builds the whole solution.
> **That objection is dissolved by the `ExcelDna.Interop` NuGet package**
> (v15.0.1, by the Excel-DNA author): it ships the Office PIAs as a
> restorable package whose own MSBuild `.targets` add the assembly
> references, so the Excel-less runner compiles against them. This is
> strictly simpler and far less fragile than hand-rolling an
> `IConnectionPoint`/`IReflect` sink with manually-transcribed `AppEvents`
> DISPIDs (the path that "can tear down the Excel host on load"). The
> reference is added in `PyExcel.Addin.csproj`; it is the *only* Office
> interop reference in the codebase.
>
> **Reference the assemblies explicitly — gotcha.** The package's own
> mechanism is a build `.targets` that flips the restored
> `lib/net452` references to `EmbedInteropTypes` after `ResolveReferences`.
> On the Windows CI runner that did **not** surface the types (the
> `lib/net452` assemblies never reached `ReferencePath`), so
> `Microsoft.Office.Interop.Excel` didn't resolve and the `Excel` alias
> fell back to our own `PyExcel.Excel` namespace —
> `error CS0234: type 'Workbook' does not exist in namespace
> 'PyExcel.Excel'`. Tried both with and without an `<EmbedInteropTypes>`
> child on the `PackageReference`; both failed identically. The fix
> (`PyExcel.Addin.csproj`): reference the package's PIA DLLs explicitly
> with `<Reference EmbedInteropTypes="true" Private="false">` pointing at
> `$(PkgExcelDna_Interop)\lib\net452\{Microsoft.Office.Interop.Excel,
> Office,Microsoft.Vbe.Interop}.dll` — the canonical way to embed a
> NuGet-delivered PIA — where `$(PkgExcelDna_Interop)` comes from
> `GeneratePathProperty="true"` on the `PackageReference`.
>
> **Second gotcha:** `GeneratePathProperty` must stand alone — do NOT pair
> it with `ExcludeAssets="all"` or `PrivateAssets="all"`, which leave the
> `$(Pkg…)` property empty (NuGet/Home #13859, #8311), so the HintPaths
> resolve to nothing and the CS0234 returns. (Cost a CI round to learn.)

### 3. CustomXMLPart persistence

**Codec half done.** `WorkbookStateCodec.Serialize` /
`WorkbookStateCodec.Deserialize` (in `src/PyExcel.State/`) handle the
XML round-trip. Schema: root `<pyexcel state-version="1"
xmlns="urn:pyexcel:state:1">` carrying `<enabled>`, optional
`<selected-script>` / `<py-input>` / `<py-output>` /
`<selected-action>`, and an `<actions>` list of `<action>` elements
(each with `name`/`script`/`input`/`output` attributes and an optional
`<kwargs>` child whose ordering is deterministic on serialise so a
no-op save doesn't churn the workbook's binary diff). Persisted
fields: `Enabled`, `SelectedScript`, `PyInput`, `PyOutput`,
`SelectedActionName`, `Actions`. Transient (NOT persisted):
`WorkbookKey` (caller supplies on Deserialize), `CurrentSheet`,
`AvailableScripts`. 15 round-trip + error-path tests in
`WorkbookStateCodecTests.cs`. Bumping the schema means bumping the
`urn:pyexcel:state:N` namespace AND
`WorkbookStateCodec.SchemaVersion`.

**COM half ✅ (code-complete — needs smoke test).** Landed as
`src/PyExcel.Addin/WorkbookStatePersister.cs` (net48, `#if NETFRAMEWORK`).
`Save(workbook, state)` deletes any existing `urn:pyexcel:state:1` parts
(`CustomXMLParts.SelectByNamespace` → `Delete`, back-to-front) then
`CustomXMLParts.Add`s the current one; `TryLoad(workbook, key)` finds the
part by namespace and returns the deserialized state (or `null`). It's a
thin shell: all XML work delegates to two new pure helpers on the tested
`WorkbookStateCodec` — `SerializeToString(state)` and
`TryDeserialize(xml, key, out state)` (the latter returns `false` for
null/blank/non-XML/foreign-namespace/wrong-version input, so a corrupt or
foreign part can't throw into the COM event handler). Both helpers are
unit-tested cross-platform (`WorkbookStatePersistenceTests`, 9 cases).
The `AppEventSink` calls `Save` on `BeforeSave`/`BeforeClose` and
`TryLoad` on `WorkbookOpen`.

### 4. UDF → kernel cancellation bridge (~80 lines)

Right now `KernelClient.Cancel(runId)` writes the CANCEL frame, the
kernel pumps it, and `pyexcel.kernel.is_cancelled()` returns `True`.
The piece **not yet wired** is the UDF firing `Cancel` when Excel-DNA
cancels its background task (formula change, workbook close).

Approach: track the live `runId` per UDF-parameter-tuple in a
`ConcurrentDictionary` inside `PyRunFunction`. When `ExcelAsyncUtil.Run`
hands us a new parameter tuple (which it does when the user changes the
formula), the previous tuple's task is no longer reachable; we'd like
to fire `Cancel(prevRunId)` then. The hook is the
`ExcelDna.Integration.ExcelAsyncUtil.Observe` overload, which gives us
the cancellation signal — switch the UDF over and call `Cancel`
when the observable disposes.

Worth it only if users hit long-running jobs in practice; Phase 4 first
slice prioritised correctness over this optimisation.

### 5. Progress UI (~200–300 lines, WinForms) — kernel half ✅, WinForms still to do

**Kernel half landed.** `pyexcel.kernel.report_progress(percent=None,
message="")` is now a user-facing helper (re-exported from
`pyexcel.kernel`, alongside `is_cancelled`). The supervisor installs a
per-job progress sink via `worker._begin_job(event, progress_sink)` and
drains it onto the wire as `PROGRESS` frames from the main loop — the
same single-writer thread that sends PONG/terminal frames, so the worker
thread never races it for the transport. Frames are flushed each
~50 ms poll tick during the run and once more after the worker joins, so
every `PROGRESS` precedes the terminal `RUN_RESULT`/`ERROR`. Meta matches
what `KernelClient.RaiseProgress` reads: `run_id`, `percent` (JSON `null`
for indeterminate updates), `message`. 5 worker unit tests + 1 e2e
supervisor test (`test_kernel_report_progress_emits_progress_frames_before_result`).
The C# side was already wired (`KernelClient.ProgressReceived` fires).

**WinForms half still to do** (needs Windows, Phase 8 charter). Wire
`KernelClient.ProgressReceived` to a small modeless WinForm with a
percent bar, message, and a Cancel button. The form lives in a new
`PyExcel.Forms` assembly but a Phase-4-grade one can be inline if you
want to move fast. The Cancel button calls `KernelClient.Cancel(runId)`
(#4's bridge is what makes that Cancel actually abort the run).

### 6. Ribbon Input/Output `{name}=Range` parser ✅

Landed on `claude/dreamy-mayer-Ud10y` — `RibbonRangeParser.Parse` in
`src/PyExcel.State/RibbonRangeParser.cs` returns an ordered list of
`RangeBinding(Name?, RangeText)` records. Handles anonymous bindings
(`A1:C10`), named bindings (`prices=A1:C10`), and the
semicolon-separated multi-binding syntax
(`prices=A1:C10; signals=D1:D10`). Empty / whitespace-only input maps
to an empty list; malformed entries (empty name, empty range,
duplicate name) throw `FormatException`. 15 unit tests in
`tests/PyExcel.Bridge.Tests/RibbonRangeParserTests.cs`.

Multi-input dispatch is done: `PyRun.ExecuteMany(script,
IReadOnlyList<object?> inputs, …)` encodes each input as its own Arrow
payload; the kernel matches them positionally to the user function's
parameters. `PyRun.Execute` (single-input) is now a thin wrapper over
it and still backs the `=PY.RUN` UDF. 5 e2e tests in `PyRunTests`.

### 7. `OnRunPython` ribbon button ✅ (code complete — needs smoke test)

Landed as `PyExcel.Excel.RangeRunner` (net48-only) +
`PyExcelRibbon.OnRunPython` delegating to it. The flow:

1. **Main thread** (ribbon callback): resolve the active workbook key,
   read `WorkbookState`, parse `PyInput` via `RibbonRangeParser`, read
   each input range off the active sheet into a 0-based `object?[,]`
   (or scalar for a single cell), capture the workbook dir for relative
   script resolution.
2. **Background `Task`**: `PyRun.ExecuteMany(...)` against
   `KernelHost.Default.Client` — the only part that blocks (SAFE-1:
   the callback has already returned).
3. **Main thread** via `ExcelAsyncUtil.QueueAsMacro`: write the decoded
   result into the `PyOutput` range. A table resizes the anchor to its
   footprint; a scalar drops into the top-left cell; a `None` return
   writes nothing.

All COM access is late-bound through `dynamic` on
`ExcelDnaUtil.Application` — **no Office PIA reference**, so the
PIA-less Windows CI build still compiles. Errors go to `Trace` +
Excel-DNA `LogDisplay`. This satisfies the Phase 4 exit criterion
("one real script runs end to end") once smoke-tested on Windows — see
the smoke-test script at the end of this doc.

### 8. `OnAddAction` / `OnEditAction` UI (Phase 8 charter)

The state plumbing already accepts what these would produce
(`StateService.AddAction` upserts by name). The form is Phase 8 work;
mention only for traceability.

---

## What's testable from CI vs. what isn't

| Component | CI-testable | Notes |
|---|---|---|
| #1 `ExcelWorkbookContext` | ✅ verified | smoke-tested on real Excel 2026-05-29 |
| #2 `AppEventSink` | ⚠️ code-complete | typed events via embedded `ExcelDna.Interop`; **compiles on Windows CI**, runtime needs the smoke test |
| #3 CustomXMLPart | ✅ codec + ⚠️ COM | codec `WorkbookStateCodec` + 24 tests (CI); `WorkbookStatePersister` compiles on Windows CI, runtime needs the smoke test |
| #4 UDF cancel bridge | ⚠️ | Async flow testable via fake ExcelAsyncUtil; the real flow needs Excel |
| #5 Progress UI | ⚠️ | kernel half (`report_progress` → `PROGRESS`) ✅ CI-tested; WinForms still manual |
| #6 Ribbon range parser | ✅ landed | `RibbonRangeParser` + 15 tests |
| #7 OnRunPython | ✅ verified | smoke-tested on real Excel 2026-05-29 — doubling, error surfacing, multi-input all pass |
| #8 Forms | ❌ | Phase 8 |

---

## Suggested order for the next session

1. ~~**#6 Ribbon range parser**~~ — done, see `RibbonRangeParser.cs`.
2. ~~**#1 ExcelWorkbookContext**~~ — done, see `ExcelWorkbookContext.cs`. Needs Windows smoke test.
3. ~~**#3 CustomXMLPart codec**~~ (codec only) — done, see `WorkbookStateCodec.cs`. COM persister still to do.
4. ~~**#7 OnRunPython**~~ — done, see `RangeRunner.cs`. Phase 4 exit
   criterion; needs the Windows smoke test below.
5. ~~**#2 AppEventSink** + COM-side CustomXMLPart persistence~~ — done,
   see `AppEventSink.cs` + `WorkbookStatePersister.cs`. The PIA-vs-late-bound
   fork is resolved (typed events via the embedded `ExcelDna.Interop` PIA;
   see §2). Both compile on the Windows CI lane; **runtime behaviour still
   needs the Windows smoke test** (§2/§3 + the new step 9 below).
6. **#4** UDF cancel bridge and **#5** progress UI — polish. #5's kernel
   half (`report_progress` → `PROGRESS` frames) is done and CI-tested;
   the WinForms dialog is the remaining (Windows-only) piece. #4 still open.

Phase 4's headline ("one real script runs end to end") is satisfied (#7,
verified on real Excel). Phase 3's "state survives close/reopen" exit
criterion is code-complete (#2/#3) and satisfied **once the smoke test
below passes**.

---

## Windows smoke test — verifying the COM pieces

These steps verify #1 (`ExcelWorkbookContext`) and #7
(`OnRunPython` / `RangeRunner`) against a real Excel. Run on a Windows
box with 64-bit Excel + the .NET SDK.

### Build the .xll

```pwsh
dotnet build PyExcel.sln --configuration Release
# The add-in lands at:
#   src/PyExcel.Addin/bin/Release/net48/PyExcel-AddIn64.xll
# Confirm the embedded/ kernel dir ships alongside it (PythonResolver
# walks up for embedded/pyexcel/kernel/__main__.py). If it's missing,
# copy the repo's embedded/ next to the .xll.
```

Make sure a Python with `pyarrow`/`pandas`/`numpy` is discoverable —
either on PATH, or point `PYEXCEL_PYTHON` at a venv's python.exe:

```pwsh
$env:PYEXCEL_PYTHON = "C:\path\to\venv\Scripts\python.exe"
```

### Load + smoke test #7 (OnRunPython end-to-end)

1. Double-click `PyExcel-AddIn64.xll` (or add it via File ▸ Options ▸
   Add-ins ▸ Manage Excel Add-ins ▸ Browse). The **PyExcel** ribbon tab
   appears.
2. Save a new workbook somewhere, e.g. `C:\tmp\smoke.xlsx`. Next to it
   create a `transform.py`:
   ```python
   def transform(df):
       return df * 2
   ```
3. In the ribbon's Python group: set **Script** to `transform.py`
   (relative — it resolves against the workbook dir), **Input** to
   `A1:B2`, **Output** to `D1`.
4. Put numbers in `A1:B2` (e.g. 1,2 / 3,4).
5. Click **Run Python**. Expected: `D1:E2` fills with 2,4 / 6,8 within
   a second or two (first click also boots the kernel, so it may take a
   beat longer).
6. **Error surfacing:** change `transform.py` to `raise ValueError("x")`,
   click Run again. The cell block is left untouched and the error
   (code/type/message/traceback) shows in Excel-DNA's LogDisplay window
   (View ▸ Log, or it auto-pops). No Excel crash, no hang.
7. **Multi-input:** Input `prices=A1:A2; signals=B1:B2`, a two-arg
   `def transform(prices, signals): return [p+s for p,s in zip(prices, signals)]`,
   Output `D1`. Confirm the column result lands.

### Smoke test #1 (workbook identity)

8. Open a *second* workbook. Set its Script/Input/Output differently.
   Switch back and forth between the two windows — each should show its
   own ribbon values (this exercises `ExcelWorkbookContext` +
   `StateService` keying). With #2 (`AppEventSink`) now wired, the ribbon
   should **auto-refresh on the switch itself** (`WorkbookActivate` →
   `RequestRibbonInvalidate`); if it doesn't, that's the spot to debug.

### Smoke test #2/#3 (state survives save → close → reopen)

9. In the first workbook: Enable PyExcel, set Script/Input/Output, and
   **save** the workbook (`.xlsx` is fine — CustomXMLParts persist in the
   modern formats). Close it. Reopen it. Expected: the ribbon comes back
   Enabled with the same Script/Input/Output you set — `WorkbookBeforeSave`
   wrote a `urn:pyexcel:state:1` CustomXMLPart and `WorkbookOpen` restored
   it. To confirm the part directly: rename a copy to `.zip` and look for
   `customXml/item1.xml` (or inspect via the Developer ▸ XML tools). A
   *second* save should not churn the part (serialisation is deterministic).

Report back what happens at steps 5/6/7/8/9 — especially any COM quirks
in the range read/write (e.g. a transposed result, an off-by-one
anchor, or a type that didn't round-trip). Those are the spots I
couldn't verify from Linux.

---

## What's already there for you

| Need this? | Look here |
|---|---|
| `StateService` API | `src/PyExcel.State/StateService.cs` + `StateServiceTests.cs` |
| `WorkbookState` shape | `src/PyExcel.State/WorkbookState.cs` |
| Ribbon callback wiring | `src/PyExcel.Ribbon/PyExcelRibbon.cs` (every getter pulls from `ActiveState()`) |
| Service locator | `src/PyExcel.State/PyExcelServices.cs` |
| Script directory tracking | `src/PyExcel.State/ScriptDirectoryWatcher.cs` + tests |
| `=PY.RUN` UDF wrapper | `src/PyExcel.Excel/PyRunFunction.cs` (the file you'll modify for #4) |
| Dispatch core | `src/PyExcel.Excel/PyRun.cs` (the file you may extend for #6) |
| Kernel `is_cancelled` / `report_progress` | `embedded/pyexcel/kernel/__init__.py` (re-exported from `worker`); supervisor drains the progress sink in `_run_with_cancellation` |
| Add-in entry point | `src/PyExcel.Addin/AddIn.cs` (where `AutoOpen` wires services) |
| Roadmap | [`ROADMAP.md`](../ROADMAP.md) Phase 3 + Phase 4 sections |
