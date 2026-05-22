# PyExcel — Codebase Architecture

> **Purpose of this file.** A map of the repository so a contributor (or a future
> chat) can find, edit, debug, and extend the code without re-discovering the
> layout each time. For *what to build next*, see [`ROADMAP.md`](ROADMAP.md).

## What PyExcel is

A Windows Excel add-in that runs user-authored Python `transform()` functions
against workbook ranges. The user selects input ranges, picks a script, clicks
**Run**; the add-in marshals the ranges to the Python side, executes the script
in an isolated project virtual environment, and writes the returned tables,
lists, scalars, and charts back into Excel. Runs are non-reactive: nothing
executes until **Run** is clicked.

## Two codebases, one product

The repository contains **two implementations** in parallel:

| | v1 — current | v2 — target |
| --- | --- | --- |
| Add-in | `PyExcel.xlam` (VBA) | `PyExcel-AddIn64.xll` (.NET / Excel-DNA) |
| Source | `src/module/*.bas`, `*.cls`, `*.frm` | `src/PyExcel.*/` (C#) |
| Python runtime | `src/embedded/*.py` | `embedded/pyexcel/` (kernel package) |
| Excel ↔ Python IPC | typed XML files on disk + `meta.xml` polling | named pipe + binary framing + Apache Arrow |
| Status | **ships today**, but carries significant tech debt | **~5% built** (Phase 1 skeleton only) |

v2 is a deliberate re-architecture, not a line-by-line port — it deletes whole
classes of v1 bug (see [`ROADMAP.md`](ROADMAP.md) → *Bug disposition*). **v1
remains the shipping product until v2 reaches parity.** Do not ship a hybrid.

⚠️ **Easy to confuse — read carefully:**
- `src/embedded/` = v1 Python (`tools.py`, `xmlParsing.py`). `embedded/pyexcel/` = v2 Python kernel.
- `src/Ribbon/RibbonUI.xml` = v1 ribbon. `src/PyExcel.Ribbon/Resources/RibbonUI.xml` = v2 ribbon.

## Repository map

```
PyExcel/
├── README.md                     Product/user documentation (describes v1)
├── ARCHITECTURE.md               This file — codebase map
├── ROADMAP.md                    Phased task list / path to production
├── PyExcel.xlam                  v1 compiled add-in (binary)
├── requirements.txt              v1-era dependency list (155 pkgs — see ROADMAP, to be slimmed)
├── PyExcel.sln                   v2 .NET solution
├── Directory.Build.props         v2 shared MSBuild settings (version, lang, warnings)
├── .gitignore
│
├── src/                          ── v1 SOURCE (VBA) and v2 SOURCE (.NET) ──
│   ├── module/                   v1: VBA modules under version control
│   │   ├── Setup.bas               First-run wizard (folders, venv, extract, pip)
│   │   ├── Update.bas              Version check, SmartClean, dependency sync
│   │   ├── HostManager.bas         Workbook registry, watchdog, ribbon state, logging
│   │   ├── CAppEvents.cls          Application event sink
│   │   ├── python.bas              Run orchestrator (serialize → spawn → poll → paste)
│   │   ├── pythonUtils.bas         Meta/artifact parsing, paste to ranges, archive
│   │   ├── xmlParsing.bas          Range ↔ typed XML
│   │   ├── chartBuilder.bas        Plotly/Matplotlib XML → native Excel chart
│   │   ├── Import.bas / Export.bas / Paste.bas
│   │   ├── modRibbon.bas           Ribbon callbacks for every button
│   │   ├── PathUtils.bas / modDst.bas
│   │   └── frm*.frm / *.frx        9 UserForm dialogs (+ frmEditActionOLD — dead)
│   ├── embedded/                 v1: Python runtime extracted into the project on setup
│   │   ├── tools.py                run_script_cli, artifact materializer
│   │   ├── xmlParsing.py           read_xml/write_xml, ExcelFormula, Plotly→XML
│   │   ├── requirements.txt        v1 project venv deps (39 pkgs)
│   │   └── instructions.txt / Uninstall.txt
│   ├── Ribbon/                   v1: RibbonUI.xml + customLogo.png
│   ├── PyExcel.Common/           v2: logging + shared types (net48 + netstandard2.0)
│   │   └── Logging/                ILog, FileLog, NullLog
│   ├── PyExcel.Ribbon/           v2: ExcelRibbon subclass + Resources/RibbonUI.xml (net48)
│   └── PyExcel.Addin/            v2: .xll entry point — AddIn.cs + PyExcel-AddIn.dna (net48)
│
├── embedded/                     ── v2 PYTHON KERNEL (shipped inside the .xll) ──
│   └── pyexcel/
│       ├── __init__.py
│       └── kernel/
│           └── framing.py          Wire framing protocol (done + tested)
│
├── tests/                        ── v2 TESTS (pytest, cross-platform) ──
│   ├── conftest.py                 Puts embedded/ on sys.path
│   └── kernel/test_framing.py
│
└── docs/
    └── v2-build.md               v2 build commands + Phase 1 exit gate
```

## v1 architecture (current product)

**Components.** `modRibbon.bas` handles ribbon clicks → `python.bas` orchestrates a
run → `xmlParsing.bas` serializes ranges → a spawned `python.exe` runs
`src/embedded/tools.py` → `pythonUtils.bas` + `chartBuilder.bas` paste results
back. `HostManager.bas` + `CAppEvents.cls` track which workbook is active and
keep the ribbon state in sync. `Setup.bas` / `Update.bas` provision and update
the project venv.

**How a run works:**
1. The **Input** ribbon field is parsed; ranges are serialized to
   `Temp/in_<script>_<runid>.xml`.
2. `python.bas` shells `python.exe userScripts/<script>.py --in … --out … --meta … --run-id …`.
3. VBA **polls `meta.xml`** on disk for `status = done | error | in_progress`,
   watching a heartbeat timestamp to detect stalls.
4. On `done`, artifacts are read from `Temp/` and pasted to the mapped ranges.
5. The whole run is copied to `Archive/` (last 10 kept).

Key property: **file-based IPC + polling.** This is the source of most v1
defects (locale-dependent number text, date-epoch guessing, polling hangs,
orphaned processes). v2 replaces it wholesale.

## v2 architecture (target)

**Process model.** Excel loads the `.xll` once (`AddIn.AutoOpen`). A single
long-lived **Python kernel** subprocess is spawned lazily and supervised; the
C# side talks to it over a **named pipe**.

**Projects** (most do not exist yet — see [`ROADMAP.md`](ROADMAP.md)):

| Project | Owns | Status |
| --- | --- | --- |
| `PyExcel.Common` | Logging (`ILog`), shared types | Phase 1 ✅ |
| `PyExcel.Ribbon` | Ribbon callbacks, `RibbonUI.xml` | Phase 1 skeleton ✅ |
| `PyExcel.Addin` | `.xll` entry point, service lifetime | Phase 1 skeleton ✅ |
| `PyExcel.Bridge` | Frame transport, named pipe, `KernelSupervisor` | Phase 2 |
| `PyExcel.Kernel.Client` | Typed run/cancel/progress API over frames | Phase 2 |
| `PyExcel.State` | Workbook registry, enabled state, ribbon state | Phase 3 |
| `PyExcel.Excel` | Range ↔ Arrow marshalling, paste, import/export | Phase 4–5 |
| `PyExcel.ChartBuilder` | Chart spec → native Excel chart | Phase 6 |
| `PyExcel.Setup` | venv, pip, project provisioning | Phase 7 |
| `PyExcel.Forms` | WinForms dialogs | Phase 8 |
| `pyexcel.kernel` (Python) | framing, supervisor, worker, transform runner | Phase 2+ |

**How a run will work:**
1. `OnRunPython` (ribbon) parses the Input/Output fields and **enqueues** a job to
   a background service, then returns immediately (rule **SAFE-1**: ribbon
   callbacks never block on the pipe).
2. `PyExcel.Excel` marshals the input ranges into an **Arrow IPC stream**.
3. `PyExcel.Kernel.Client` sends a `RUN_REQUEST` frame; the kernel runs the
   user `transform()` and streams `PROGRESS` then `RUN_RESULT` frames back.
4. `PyExcel.Excel` writes the Arrow result tables to the mapped ranges;
   `PyExcel.ChartBuilder` builds any charts.

### The kernel protocol

`embedded/pyexcel/kernel/framing.py` defines the wire format: a length-prefixed
binary frame = `body_len · type · meta_json · payloads`. Frame types
(`HELLO`, `PING`/`PONG`, `RUN_REQUEST`/`RUN_RESULT`, `PROGRESS`, `LOG`,
`CANCEL`, `ERROR`, …) are a stable enum. `PROTOCOL_VERSION = 2`. The C#
counterpart (`PyExcel.Bridge/Framing.cs`) **must match this byte-for-byte** —
build cross-language conformance tests. Frame sizes are bounded (256 MiB) on
both encode and decode to bound memory against a malformed peer.

## Where to make changes (navigation)

| To change… | v1 (today) | v2 (target) |
| --- | --- | --- |
| A ribbon button's behaviour | `modRibbon.bas` | `PyExcel.Ribbon/PyExcelRibbon.cs` |
| Ribbon layout (tabs/groups) | `src/Ribbon/RibbonUI.xml` | `src/PyExcel.Ribbon/Resources/RibbonUI.xml` |
| How a run is orchestrated | `python.bas` | `PyExcel.Kernel.Client` + `PyExcel.Bridge` |
| Range → Python data marshalling | `xmlParsing.bas` + `xmlParsing.py` | `PyExcel.Excel` + `pyexcel.kernel` (Arrow) |
| Writing results back to ranges | `pythonUtils.bas` | `PyExcel.Excel` |
| Chart rendering | `chartBuilder.bas` | `PyExcel.ChartBuilder` |
| The `transform()` contract / runner | `src/embedded/tools.py` | `embedded/pyexcel/kernel/worker.py` |
| The wire protocol | n/a | `embedded/pyexcel/kernel/framing.py` ↔ `PyExcel.Bridge/Framing.cs` |
| Import / Export / Paste | `Import.bas` / `Export.bas` / `Paste.bas` | `PyExcel.Excel` services |
| First-run setup / venv | `Setup.bas` | `PyExcel.Setup` |
| Dialogs | `frm*.frm` | `PyExcel.Forms` |
| Workbook/ribbon state | `HostManager.bas`, `CAppEvents.cls` | `PyExcel.State` |
| Logging | `HostManager.bas` `LogToFile` | `PyExcel.Common/Logging/FileLog.cs` |
| Project dependencies | `src/embedded/requirements.txt` | same, slimmed (see ROADMAP) |

## Build & test

Cross-platform (Linux/macOS — what CI should run):
```bash
dotnet build src/PyExcel.Common/PyExcel.Common.csproj --framework netstandard2.0 -c Release
pytest tests/
```
Windows (produces the `.xll`):
```powershell
dotnet restore PyExcel.sln
dotnet build PyExcel.sln -c Release --no-restore
```
Full build prerequisites and the Phase 1 exit gate are in [`docs/v2-build.md`](docs/v2-build.md).
There is **no CI pipeline yet** — adding one is a cross-cutting task in `ROADMAP.md`.

## Conventions

- **Logging.** Both v1 and v2 write to `%TEMP%\PyExcel_Debug.log`, format
  `[YYYY-MM-DD hh:mm:ss.fff] [LEVEL] message`. v2 uses `ILog` (`PyExcel.Common`).
- **Versioning.** v2 .NET assemblies: SemVer `2.0.0-alpha` (`Directory.Build.props`).
  v2 Python kernel: PEP 440 `2.0.0a0` (`embedded/pyexcel/__init__.py`) — the same
  release in each ecosystem's convention. v1 uses a build timestamp.
- **Phase model.** Phases 1–9 are defined in [`ROADMAP.md`](ROADMAP.md) — the
  single source of truth. `// PHASE n` comments in C# must match it.
- **SAFE-1.** Ribbon callbacks must never block on the kernel pipe — enqueue and return.
- **Branching.** Feature work happens on a dedicated branch, not `main`.

## Known issues & tech debt

The full audit — bugs, dead code, inefficiencies, UX gaps — is folded into
[`ROADMAP.md`](ROADMAP.md) under *Bug disposition* and *Cut list*, classified by
whether the v2 architecture deletes them, whether they must be fixed during the
port, or whether they need fixing in v1 now.
