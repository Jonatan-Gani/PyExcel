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
- [ ] **Strip personal data** — delete the `if __name__ == "__main__"` block (~570 lines, hardcoded author coursework path) from `src/embedded/xmlParsing.py`.
- [ ] **Slim dependencies** — replace the 155-package root `requirements.txt` (jupyterlab, pygame, yt-dlp, telegram-bot, Flask, …) with the minimal set the kernel needs; resolve it against `src/embedded/requirements.txt` into one canonical UTF-8 list. Add `pyarrow` (v2 IPC needs it). Audit the two `git+https` deps.
- [ ] **Decide v1 maintenance policy** — confirm v1 stays frozen (security fixes only) while v2 is built; no new v1 features.
- [ ] Add a CI workflow (see *Cross-cutting* below) so every subsequent phase is gated.

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

- [ ] `PyExcel.Bridge/Framing.cs` — mirror `framing.py` byte-for-byte (same frame layout, bounds, determinism).
- [ ] Cross-language conformance tests: frames encoded in C# decode in Python and vice versa.
- [ ] Named-pipe transport with a SID/ACL check (reject non-owner connections → `ERROR` frame).
- [ ] `KernelSupervisor` — spawn `python -m pyexcel.kernel` with an **argv array** (no `cmd /c` string), `HELLO` handshake, `PING`/`PONG` health checks, deterministic kill on `AutoClose` / crash / hang.
- [ ] `PyExcel.Kernel.Client` — typed API: `RunRequest`, `RunResult`, `Progress`, `Cancel`, `Log` over frames.
- [ ] Python `supervisor.py` — accept the pipe connection, dispatch frames to workers.
- [ ] Python `worker.py` — run one job: receive request → execute → reply.
- [ ] Python `arrow_io.py` — DataFrame ↔ Arrow IPC stream.
- [ ] Bounded/malformed-frame handling on the C# side (mirror the `framing.py` test suite).

**Exit criteria.** C# spawns the kernel, completes a `HELLO`/`PING`/`PONG`
round-trip, and kills it cleanly on shutdown. Framing conformance tests pass in
both languages. No orphaned `python.exe` after Excel closes.

---

## Phase 3 — State & events

**Objective.** The ribbon accurately reflects per-workbook state; switching
sheets/workbooks never shows stale values.

**Deliverables.** `PyExcel.State`, an `AppEventSink`.

- [ ] `StateService` — per-workbook enabled flag, current sheet, host-workbook registry. One source of truth (no module globals).
- [ ] Persist per-workbook state (decide storage — see *Open decisions*).
- [ ] `AppEventSink` — `WorkbookOpen`/`Activate`/`SheetActivate` → update state + invalidate ribbon.
- [ ] Wire `RibbonEnabled` / all `getEnabled` to `StateService` (replace the hardcoded `false`).
- [ ] `FileSystemWatcher` on `userScripts/` → refresh the Script dropdown.
- [ ] Wire `OnAddAction` / `OnEditAction` / `OnDeleteAction` and the action/script/input/output getters to real state.

**Exit criteria.** Enabling a workbook lights the ribbon; switching sheet or
workbook updates every ribbon field correctly; state survives a save/reopen.

---

## Phase 4 — Excel marshalling & first run (the thin slice)

**Objective.** Prove one full run works end to end — this de-risks the pipe,
Arrow, COM interop, and the threading model in a single slice before going wide.

**Deliverables.** `PyExcel.Excel`; a wired `OnRunPython`.

- [ ] Range → Arrow: read range `Value2` into a typed Arrow record batch (handle dates, blanks/`NA`, Excel error values).
- [ ] Arrow → Range: write result tables, lists, and scalars to ranges.
- [ ] Parse the Input/Output ribbon fields, including the `{name}=Range` syntax.
- [ ] Wire `OnRunPython`: parse → marshal → enqueue to `Kernel.Client` → write results. **SAFE-1**: enqueue and return; never block the callback.
- [ ] Non-blocking progress UI driven by `PROGRESS` frames, with a working **Cancel**.
- [ ] Archive a run (inputs, outputs, log); retention cap.
- [ ] Surface kernel errors (full traceback, script name, log path) clearly to the user.
- [ ] `ExcelFormula` round-trip (A1 mode).
- [ ] Tests: range ↔ Arrow round-trip incl. dates, nulls, formulas.

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

Typosquat removal ✅ · UTF-16 requirements ✅ · dependency slimming · personal-data removal.

## Cut list — delete, do not migrate (~3,000+ lines)

- [ ] `xmlParsing.bas` — entire file (Arrow replaces it).
- [ ] `chartBuilder.bas` lines ~2–1010 — dead alternate engine.
- [ ] `tools.py` — ~1,200 commented-out legacy lines.
- [ ] `xmlParsing.py` — ~485 commented lines + the ~570-line `__main__` block.
- [ ] `frmEditActionOLD.frm/.frx`; commented blocks in the three `frmEdit*` forms; the `CAppEvents.cls` duplicate.
- [ ] `Import.bas ReadExcel_ADO` (~100 dead lines); unused helpers across modules.
- [ ] Most of `Update.bas` (SmartClean / manifest / version-name machinery).
- [ ] `PyExcel.xlam` and all of `src/module/` — at Phase 9.

## Cross-cutting — CI & testing

- [ ] Add `.github/workflows/` — on every push: build the `PyExcel.Common` netstandard2.0 slice + run `pytest tests/` (Linux); build the full solution on a Windows runner.
- [ ] Gate merges on green CI.
- [ ] Every phase lands with tests; track coverage of `PyExcel.Excel`, `PyExcel.Bridge`, and the kernel.
- [ ] Code review on every PR against the *Definition of production-grade*.

## Open decisions

1. **Chart transport** — kernel emits a JSON chart spec in the `RUN_RESULT` frame, `PyExcel.ChartBuilder` builds the Excel chart? *(Recommended — consistent with the framing protocol.)*
2. **Lists / scalars** — carry as small Arrow batches, or as JSON in the frame `meta`? *(JSON is simpler for tiny payloads.)*
3. **Per-workbook state storage** — defined names, or a hidden sheet? Affects Phase 3 and the Phase 9 migration.
4. **`PyExcel.Forms` UI tech** — WinForms (lower friction with Excel-DNA, matches existing layouts) vs WPF.
