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

### 2. `AppEventSink` — the Excel.Application event subscriber (~150 lines)

Subscribes to `Application.WorkbookOpen`, `WorkbookActivate`,
`WorkbookBeforeClose`, `SheetActivate`. Each handler:

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

> **⚠️ Design decision before writing this — PIA vs. late-bound.**
> Subscribing to `Application.WorkbookOpen` &c. the easy way means a
> typed `Microsoft.Office.Interop.Excel` reference and `app.WorkbookOpen
> += handler`. **But that breaks the Windows CI build:** GitHub's
> `windows-latest` runner has no Excel installed, so the Office PIA
> isn't resolvable, and the CI lane builds the *whole* solution
> (`dotnet build PyExcel.sln`). Everything COM-bound so far
> (`ExcelWorkbookContext`, `RangeRunner`) deliberately uses late-bound
> `dynamic` on `ExcelDnaUtil.Application` for exactly this reason — no
> PIA, CI stays green. An event sink can't use `dynamic +=`, so it has
> to advise the connection point manually (`IConnectionPointContainer`
> / `IConnectionPoint` against the `AppEvents` dispinterface) — doable
> but fiddly and untestable from CI, and a malformed sink can tear down
> the Excel host on load. The alternatives are: (a) write the
> late-bound `IConnectionPoint` sink; (b) take the PIA reference with
> `EmbedInteropTypes=true` and *exclude `PyExcel.Addin` from the CI
> solution build* (build it only in a separate Excel-present job, or
> drop it from `PyExcel.sln`'s CI configuration). This is a real fork
> in the road — pick deliberately rather than reflexively.

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

**COM half still to do:** Windows-only `WorkbookStatePersister` that
on `WorkbookBeforeSave` calls `Serialize`, writes the result into the
workbook's `CustomXMLPart` collection (deleting any existing
`urn:pyexcel:state:1` part first), and on `WorkbookOpen` finds the
part by namespace and calls `Deserialize` back into the
`StateService`. Lives in `PyExcel.Addin` under `#if NETFRAMEWORK`.

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

### 5. Progress UI (~200–300 lines, WinForms)

Wire `KernelClient.ProgressReceived` to a small modeless WinForm with a
percent bar, message, and a Cancel button. The form lives in a new
`PyExcel.Forms` assembly (Phase 8 charter) but a Phase-4-grade one
can be inline if you want to move fast. The Cancel button calls
`KernelClient.Cancel(runId)`.

The kernel currently emits no `PROGRESS` frames. Add a
`pyexcel.kernel.report_progress(percent, message)` helper user scripts
can call; supervisor sends it as `PROGRESS` over the wire. C# side is
already wired (`KernelClient.ProgressReceived` event fires).

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
| #1 `ExcelWorkbookContext` | ✅ landed | COM-bound; manual smoke test still required |
| #2 `AppEventSink` | ❌ | COM events; manual smoke test |
| #3 CustomXMLPart codec | ✅ landed | `WorkbookStateCodec` + 15 tests; COM read/write part still to do |
| #4 UDF cancel bridge | ⚠️ | Async flow testable via fake ExcelAsyncUtil; the real flow needs Excel |
| #5 Progress UI | ❌ | WinForms; manual |
| #6 Ribbon range parser | ✅ landed | `RibbonRangeParser` + 15 tests |
| #7 OnRunPython | ✅ code complete | `RangeRunner` + ExecuteMany tested; range read/write needs Windows smoke test |
| #8 Forms | ❌ | Phase 8 |

---

## Suggested order for the next session

1. ~~**#6 Ribbon range parser**~~ — done, see `RibbonRangeParser.cs`.
2. ~~**#1 ExcelWorkbookContext**~~ — done, see `ExcelWorkbookContext.cs`. Needs Windows smoke test.
3. ~~**#3 CustomXMLPart codec**~~ (codec only) — done, see `WorkbookStateCodec.cs`. COM persister still to do.
4. ~~**#7 OnRunPython**~~ — done, see `RangeRunner.cs`. Phase 4 exit
   criterion; needs the Windows smoke test below.
5. **#2 AppEventSink** + COM-side CustomXMLPart persistence. **Blocked
   on a design decision** — see the note under §2: a typed Office PIA
   reference is the easy path but breaks the PIA-less Windows CI build,
   so the event sink must wire COM events late-bound. Still open.
6. **#4** UDF cancel bridge and **#5** progress UI — polish, still open.

Phase 4's headline ("one real script runs end to end") is satisfiable
once the #7 smoke test below passes. Phase 3's exit criteria still want
#2 (state surviving close/reopen).

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
   `StateService` keying). NB: until #2 (`AppEventSink`) lands, the
   ribbon may not auto-refresh on the *switch* itself — toggle a field
   or re-click the tab to force a redraw. Auto-refresh-on-activate is
   exactly what #2 adds.

Report back what happens at steps 5/6/7/8 — especially any COM quirks
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
| Kernel `is_cancelled` | `embedded/pyexcel/kernel/__init__.py` (re-exported from `worker`) |
| Add-in entry point | `src/PyExcel.Addin/AddIn.cs` (where `AutoOpen` wires services) |
| Roadmap | [`ROADMAP.md`](../ROADMAP.md) Phase 3 + Phase 4 sections |
