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

**Current position:** Phase 0 ✅ · Phase 1 ✅ · Phase 2 ✅ · **Phase 3 / Phase 4 next.** A focused brief for the next session — what shipped, the public APIs, what Phase 4 still has to build — is in [`docs/phase4-handoff.md`](docs/phase4-handoff.md).

Terse running record (newest first) so a new session can pick up where work stopped:

- **2026-05-23 — Phase 2 complete.** Shipped the rest of the kernel data plane in one session, both CI lanes green on `ae0b4f0`:
  - `arrow_io.py` — shape-preserving Arrow IPC for DataFrame / Series / list / tuple / 1-D-or-2-D numpy / scalar, with `pyexcel-shape` and `pyexcel-orientation` schema metadata so the host can reconstruct cell geometry. 39 pytest tests.
  - `worker.py` — pure `run_job(meta, payloads) -> JobOutcome`; loads the user script (mtime-cached), decodes Arrow payloads, calls the target function, replies with `RUN_RESULT` or a typed `ERROR` (8 stable codes from `BadRequest` through `Exception`). 23 unit tests + 2 e2e supervisor tests.
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

- [ ] `StateService` — per-workbook enabled flag, current sheet, host-workbook registry. One source of truth (no module globals).
- [ ] Persist per-workbook state as a `CustomXMLPart` on the workbook.
- [ ] `AppEventSink` — `WorkbookOpen`/`Activate`/`SheetActivate` → update state + invalidate ribbon.
- [ ] Wire `RibbonEnabled` / all `getEnabled` to `StateService` (replace the hardcoded `false`).
- [ ] `FileSystemWatcher` on `userScripts/` → refresh the Script dropdown.
- [ ] Wire `OnAddAction` / `OnEditAction` / `OnDeleteAction` and the action/script/input/output getters to real state.

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
