# PyExcel — VBA → .NET Migration Audit & Roadmap

**Date:** 2026-05-22
**Goal:** Move as much as possible out of the v1 VBA add-in into the v2 .NET / Excel-DNA add-in.

## Target end state

| Layer | Today (v1) | After migration |
| --- | --- | --- |
| Orchestration, Excel I/O, UI, setup | ~12k lines VBA across 20 `.bas`/`.cls`/`.frm` | **0 lines VBA** — all .NET (`src/PyExcel.*`) |
| Data interchange | Typed XML files on disk + `meta.xml` polling | Named-pipe kernel + binary framing + Arrow IPC |
| Python | `tools.py` + `xmlParsing.py` (~3.7k lines, ~60% dead) | One `pyexcel.kernel` package (~6 small modules) |
| Distribution | `PyExcel.xlam` | `PyExcel-AddIn64.xll` |

"As much as possible" realistically means **everything except the Python kernel** — the kernel must stay Python because it runs user `transform()` code. Everything else moves to .NET.

## Assumptions (correct me if wrong)

1. v2's already-chosen architecture stands: a **persistent Python kernel** spoken to over a **named pipe** with the `framing.py` protocol + **Apache Arrow** payloads — not the v1 file-based XML IPC.
2. v1 (`.xlam`) remains the shipping product **until** v2 reaches feature parity; we do not ship a half-ported hybrid.
3. The 9 VBA UserForms are rewritten as .NET WinForms (the `.frx` binary resources cannot be ported).
4. This is a multi-month effort; the table below is relative effort (S/M/L/XL), not a schedule.

---

## Why this is a rewrite, not a line-by-line port

The migration is worth doing because the v2 architecture **deletes whole bug classes for free**:

- **Arrow replaces typed XML.** Arrow carries column types, nulls, and timestamps natively. This deletes `xmlParsing.bas` (1,794 lines) and the `read_xml`/`write_xml` halves of the Python files — and with them the locale-dependent number formatting, `CLng` overflow, Excel-epoch date guessing, and type-inference bugs. None of that gets ported.
- **The framing protocol replaces `meta.xml` polling.** `PROGRESS`/`RUN_RESULT` frames and clean EOF semantics delete the v1 poll loop, the heartbeat false-stall logic, and the run-id-mismatch hang.
- **A `KernelSupervisor` owns process lifetime.** This deletes the orphaned-subprocess bug — the kernel is one long-lived child, spawned with a real argument array (no `cmd /c` string), killed deterministically.
- **`.xll` replacement is the update mechanism.** Excel-DNA reloads the new `.xll` on next launch. This deletes most of `Update.bas` (1,037 lines): SmartClean, version-name bookkeeping, manifest diffing.
- **.NET embedded resources replace the base64 `EmbeddedStore` sheet.** This deletes the chunked-base64 extraction in both `Setup.bas` and `Update.bas` — including the multi-chunk gap bug that silently corrupts files.
- **Real exceptions + `ILog` replace `On Error Resume Next`.** The silent-failure pattern stops being the default.

So the bulk of v1 isn't ported — it's **retired**. Roughly half the VBA disappears rather than moving.

---

## Component migration map

| v1 component | ~Lines | → Destination | Disposition | Effort |
| --- | --- | --- | --- | --- |
| `modRibbon.bas` | 1,956 | `PyExcel.Ribbon` (.NET) | **Port** callbacks; skeleton already exists. Drop the stale `currentSheetName` global — read state from `PyExcel.State`. | M |
| `python.bas` | 673 | `PyExcel.Bridge` + `PyExcel.Kernel.Client` | **Rewrite.** Spawn/poll → pipe client. Logic mostly disappears. | M |
| `pythonUtils.bas` | 1,063 | `PyExcel.Excel` | **Port** the paste/archive logic; artifact reading becomes "write Arrow result tables to ranges". | M |
| `xmlParsing.bas` | 1,794 | — | **DELETE.** Replaced by Arrow marshalling in `PyExcel.Excel`. ~70% is already dead code. | — |
| `chartBuilder.bas` | 2,031 | `PyExcel.ChartBuilder` (.NET) | **Port** the live ~1,000 lines of Excel-chart COM; drop the dead ~1,000. Reads a chart spec (JSON), not XML. | L |
| `Setup.bas` | 1,523 | `PyExcel.Setup` (.NET) | **Port** venv creation + pip install; resource extraction → .NET embedded resources. | M |
| `Update.bas` | 1,037 | mostly — | **DELETE.** Keep only venv dependency-sync → `PyExcel.Setup`. | S |
| `HostManager.bas` | 408 | `PyExcel.State` (.NET) | **Port** workbook registry / ribbon-enabled state. | S–M |
| `CAppEvents.cls` | 81 | `PyExcel.Excel` (`AppEventSink`) | **Port** — already planned in `AddIn.cs` comments. | S |
| `Import.bas` | 505 | `PyExcel.Excel` (import service) | **Port**, but use a real CSV library — do not port the naive parser. | M |
| `Export.bas` | 77 | `PyExcel.Excel` (export service) | **Port**; fix RFC-4180 quoting in transit. | S |
| `Paste.bas` | 130 | `PyExcel.Excel` | **Port.** | S |
| `PathUtils.bas` | 140 | `PyExcel.Common` / `PyExcel.Setup` | **Port**; fix UNC + localized-SharePoint resolution in transit. | S |
| `modDst.bas` | 307 | `PyExcel.Excel` | **Port** range resolution/formatting. | S |
| 9 × `frm*.frm` | ~2,900 | `PyExcel.Forms` (.NET WinForms) | **Rewrite.** `.frx` resources cannot be ported. | L |
| `frmEditActionOLD.frm/.frx` | — | — | **DELETE** — superseded. | — |
| `tools.py` | 1,748 | `pyexcel.kernel.worker` | **Rewrite/restructure.** Keep transform-running + artifact materialization; the XML writers die (Arrow). ~70% is dead code. | M |
| `xmlParsing.py` | 1,942 | `pyexcel.kernel.figures` + `formula` | `read_xml`/`write_xml` **DELETE** (Arrow). Keep `PlotlyToExcelXMLConverter` traversal (emit a JSON spec, not XML) and `ExcelFormula`. | M |
| `PyExcel.xlam` | — | — | **DELETE** — replaced by the `.xll`. | — |

---

## Current v2 state (Phase 1 only — ~5% done)

Done: `PyExcel.Common` (logging/types), `PyExcel.Ribbon` skeleton, `PyExcel.Addin` skeleton, `pyexcel/kernel/framing.py` (well-built, tested).

Not started: every ribbon callback is a `StubAction`; `RibbonEnabled` always returns `false`; no bridge, no kernel supervisor/worker, no Excel interop, no state service, no setup, no forms. `embedded/pyexcel/kernel/` contains only `framing.py`.

---

## Bug disposition

Carried from the prior audit, reclassified for the migration.

### Designed out — do NOT port (the new architecture removes them)

- [ ] Locale-dependent `CStr(CDbl())`/`Format$` numeric corruption — Arrow.
- [ ] `CLng` overflow on 10-digit integers — Arrow.
- [ ] Excel-epoch / timestamp-heuristic date bugs — Arrow timestamp type.
- [ ] Column/scalar type-inference mistakes — explicit Arrow schema.
- [ ] `meta.xml` polling hang, heartbeat false-stall, run-id mismatch — framing protocol.
- [ ] Orphaned `python.exe` after timeout — `KernelSupervisor` lifetime.
- [ ] `cmd /c` command-line quoting/injection — kernel spawned with an argv array.
- [ ] Multi-chunk base64 extraction gap bug — .NET embedded resources.
- [ ] UTF-16 `requirements.txt` that pip can't read — .NET writes UTF-8.
- [ ] `xmlParsing.bas` last-row format clobber — paste rewritten on Arrow.

### Fix in transit — must be done deliberately during the port

- [ ] **CSV parsing** (`Import.bas`, `Export.bas`, `Paste.bas`, `chartBuilder` CSV helpers) — adopt a real RFC-4180 library; handle UTF-8/BOM, embedded newlines, TSV.
- [ ] **Silent failure** — `On Error Resume Next` + `Debug.Print` becomes `ILog` + a surfaced error. This is discipline, not free; enforce it in review.
- [ ] **Stale sheet state** — redesign as a single source of truth in `PyExcel.State`; no module globals.
- [ ] **Subprocess hardening** — `KernelSupervisor` must kill the kernel on `AutoClose`, on crash, and on Excel hang.
- [ ] **frmExportWizard dead buttons** — the WinForms rewrite must actually wire the row edit/remove handlers.
- [ ] **chartBuilder orphan charts / null-attribute crashes** — add proper guards in `PyExcel.ChartBuilder`.
- [ ] **Destructive paste with no confirmation** — UX decision; carries over unless changed.
- [ ] **`Setup` diagnostics** — surface kernel/venv/pip stdout+stderr to the user instead of discarding it.

### Fix now — independent of the migration (v1 still ships meanwhile)

- [ ] **Remove the typosquat** `panadas==0.2` from `requirements.txt:84` — supply-chain risk, live today.
- [ ] **Slim dependencies** — the 155-package `requirements.txt` (jupyterlab, pygame, yt-dlp, telegram-bot, Flask…) should drop to what the kernel needs: `pandas numpy pyarrow plotly matplotlib lxml` + their transitive deps. **Add `pyarrow`** — the new IPC format needs it.
- [ ] **Resolve the two conflicting requirements files** into one canonical UTF-8 list.
- [ ] **Strip personal data** — `xmlParsing.py` hardcodes the author's coursework path and ~570 lines of private code.

---

## Proposed phased roadmap

> Note: this dependency order differs slightly from `docs/v2-build.md` (whose phase numbers also disagree with the `// PHASE n` comments in the C# code). Recommend updating that doc to match whatever order is chosen.

- [ ] **Phase 2 — Bridge + kernel core.** C# framing (mirror `framing.py` byte-for-byte), named-pipe transport with an ACL check, `KernelSupervisor` (spawn/health/kill). Python side: `supervisor.py`, `worker.py`, `arrow_io.py`, `__main__.py`. Exit: a `HELLO`/`PING`/`PONG` round-trip over the pipe, kernel killed cleanly on shutdown.
- [ ] **Phase 3 — State.** `PyExcel.State`: workbook registry, per-workbook enabled flag, ribbon `getEnabled`. Exit: enabling a workbook lights up the ribbon; state survives sheet/workbook switches.
- [ ] **Phase 4 — Excel marshalling + first real run (the thin slice, see below).** `PyExcel.Excel`: Range → Arrow, Arrow → Range, wire `OnRunPython` end-to-end. Exit: a one-table-in / one-table-out script runs and writes results.
- [ ] **Phase 5 — Import / Export / Paste services.** Real CSV library; sheet-picker; format handling.
- [ ] **Phase 6 — ChartBuilder.** `PyExcel.ChartBuilder` consumes a JSON chart spec from the kernel; port the live chart-COM code.
- [ ] **Phase 7 — Setup.** venv creation, pip install, dependency sync, diagnostics surfaced. Retire `Update.bas` (update = ship a new `.xll`).
- [ ] **Phase 8 — Forms.** Rewrite the 9 dialogs as WinForms; ship the ribbon logo PNG.
- [ ] **Phase 9 — Cutover.** Feature-parity check vs v1, v1→v2 workbook-state migration, retire `PyExcel.xlam`.

## The thin slice — de-risk before going wide

Before building breadth, prove **one run works end to end**: Bridge + minimal kernel + minimal Range↔Arrow + a wired `OnRunPython` that takes one input table and writes one output table. This exercises the pipe, the framing protocol, Arrow on both sides, COM interop, and the threading model in one go — the four hardest unknowns. Everything after it is comparatively mechanical.

## Hard parts & risks

- **COM interop & threading.** Excel-DNA callbacks arrive on COM threads; the kernel pipe must be driven off a background thread. `SAFE-1` (ribbon callbacks never block on the pipe — already noted in `PyExcelRibbon.cs`) must hold: `OnRunPython` enqueues and returns.
- **`chartBuilder.bas` port.** ~1,000 lines of intricate Excel Chart COM. Mechanical but slow and error-prone; budget accordingly.
- **Forms rewrite.** 9 dialogs, full WinForms rebuild — the `.frx` blobs give you nothing. Largest pure-UI cost.
- **Arrow ↔ Excel type fidelity.** Dates, blanks/`NA`, formulas (`ExcelFormula`), and Excel error values still need explicit handling — Arrow removes the *encoding* bugs, not the *mapping* decisions.
- **Setup is Windows-coupled.** venv provisioning, PATH discovery, the Store-Python stub — hard to test on CI; keep it isolated behind an interface.
- **No automated tests for the shipping product.** Only `framing.py` is covered. Every ported module needs tests written *as* it lands.

## Cut list — delete, do not migrate

~3,000+ lines that should never reach .NET:

- [ ] `xmlParsing.bas` — entire file (Arrow replaces it).
- [ ] `chartBuilder.bas` lines ~2–1010 — dead alternate engine.
- [ ] `tools.py` — ~1,200 commented-out legacy lines.
- [ ] `xmlParsing.py` — ~485 commented lines + the ~570-line `__main__` coursework block.
- [ ] `frmEditActionOLD.frm/.frx`; commented-out blocks in the three `frmEdit*` forms; `CAppEvents.cls` duplicate.
- [ ] `Import.bas ReadExcel_ADO` (~100 dead lines); unused helpers across modules.
- [ ] Most of `Update.bas` (SmartClean / manifest / version-name machinery).
- [ ] `PyExcel.xlam` itself, at cutover.

## Open decisions

1. **Chart transport** — kernel emits a JSON chart spec in the `RUN_RESULT` frame and `PyExcel.ChartBuilder` builds the Excel chart? (Recommended — consistent with the framing protocol.)
2. **Lists / scalars** — carry as small Arrow batches, or as JSON in the frame `meta`? (JSON is simpler for tiny payloads.)
3. **WinForms vs WPF** for `PyExcel.Forms` — WinForms is the lower-friction match for Excel-DNA and the existing dialog layouts.
4. **v1→v2 state migration** — auto-migrate the per-workbook named ranges/actions, or require a re-enable? Affects Phase 9 scope.
