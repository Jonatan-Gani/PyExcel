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

### 1. `ExcelWorkbookContext` — concrete `IWorkbookContext` (~50 lines)

`PyExcel.State.IWorkbookContext` is an interface today. Production needs an
implementation that returns the active workbook's identity from
`Application.ActiveWorkbook`. Lives in `PyExcel.Addin` (or a new
`PyExcel.State.Windows`) under `#if NETFRAMEWORK`. Key strategy: use
`Workbook.FullName` for saved workbooks, fall back to `Workbook.Name +
SessionGuid` for unsaved ones (matches the contract documented in
`IWorkbookContext`).

`PyExcel.Addin.AddIn.AutoOpen` calls:

```csharp
PyExcelServices.WorkbookContext = new ExcelWorkbookContext();
```

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

Still to do: shift `PyRun.Execute`'s signature from one input to
`params object?[] inputs` so the dispatcher can wire each
`RangeBinding` as a positional arg to the user's transform function.
That refactor is bundled with #7 (`OnRunPython`) since the ribbon
button is the only caller that benefits today — the UDF stays
single-positional.

### 7. `OnRunPython` ribbon button (~100 lines)

Once #1, #2, #6 land, this is mechanical:

```csharp
public void OnRunPython(IRibbonControl control)
{
    var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
    if (key is null) return;
    var s = PyExcelServices.State.Get(key);
    if (s.SelectedScript is null) { /* surface error */ return; }
    var inputs = RibbonRangeParser.Parse(s.PyInput);
    // resolve ranges via Excel COM, marshal via ArrowMarshal, run, write
    // output back to s.PyOutput range.
}
```

This is where the "ribbon button writes results back to a range"
behaviour finally lives. Phase 4 exit criteria is satisfied once this
runs end-to-end against a real workbook.

### 8. `OnAddAction` / `OnEditAction` UI (Phase 8 charter)

The state plumbing already accepts what these would produce
(`StateService.AddAction` upserts by name). The form is Phase 8 work;
mention only for traceability.

---

## What's testable from CI vs. what isn't

| Component | CI-testable | Notes |
|---|---|---|
| #1 `ExcelWorkbookContext` | ❌ | COM-bound; manual smoke test |
| #2 `AppEventSink` | ❌ | COM events; manual smoke test |
| #3 CustomXMLPart codec | ✅ landed | `WorkbookStateCodec` + 15 tests; COM read/write part still to do |
| #4 UDF cancel bridge | ⚠️ | Async flow testable via fake ExcelAsyncUtil; the real flow needs Excel |
| #5 Progress UI | ❌ | WinForms; manual |
| #6 Ribbon range parser | ✅ landed | `RibbonRangeParser` + 15 tests |
| #7 OnRunPython | ⚠️ | Dispatcher logic yes; range read/write via COM no |
| #8 Forms | ❌ | Phase 8 |

---

## Suggested order for the next session

1. ~~**#6 Ribbon range parser**~~ — done, see `RibbonRangeParser.cs`.
2. **#1 ExcelWorkbookContext** — small, mechanical. (~1 commit)
3. ~~**#3 CustomXMLPart codec**~~ (codec only) — done, see `WorkbookStateCodec.cs`. COM persister still to do.
4. **#2 AppEventSink** + COM-side CustomXMLPart persistence + smoke
   test in Excel. **This is the milestone that unlocks the Phase 3
   exit criteria.** (~2 commits)
5. **#7 OnRunPython** wired against the new context + parser. Phase 4
   exit criteria becomes satisfiable here. (~1 commit)
6. **#4** and **#5** as polish once the rest is real. (~2 commits)

Total budget: ~7 commits, half a session of focused work — minus the
manual Excel smoke testing which has to happen on a developer's machine
either way.

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
