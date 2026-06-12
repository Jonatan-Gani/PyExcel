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

## One codebase: the .NET rewrite

PyExcel is the **.NET / Excel-DNA** implementation (the rewrite historically
called "v2"):

| | |
| --- | --- |
| Add-in | `PyExcel-AddIn64.xll` (.NET / Excel-DNA) |
| Source | `src/PyExcel.*/` (C#) |
| Python runtime | `embedded/pyexcel/` (kernel package) |
| Excel ↔ Python IPC | named pipe + binary framing + Apache Arrow |

> **History.** PyExcel began as a VBA `.xlam` add-in (v1) that used typed XML
> files on disk plus `meta.xml` polling for Excel↔Python IPC. The current
> codebase is a deliberate re-architecture — not a line-by-line port — that
> deletes whole classes of v1 bug (see [`ROADMAP.md`](ROADMAP.md) → *Bug
> disposition*). The v1 VBA sources, the compiled `PyExcel.xlam`, and its
> on-disk Python runtime have been **removed from the repository**; what remains
> is the rewrite. It is in active development and has not yet passed a Windows +
> Excel smoke test — see [`ROADMAP.md`](ROADMAP.md) for live status.

## Repository map

```
PyExcel/
├── README.md                     Product overview
├── ARCHITECTURE.md               This file — codebase map
├── ROADMAP.md                    Phased task list / path to production
├── requirements.txt              Python kernel dependencies (the slimmed kernel set)
├── PyExcel.sln                   .NET solution
├── Directory.Build.props         Shared MSBuild settings (version, lang, warnings)
├── .gitignore
│
├── src/                          ── SOURCE (.NET / C#) ──
│   ├── PyExcel.Common/           Logging + shared types (net48 + netstandard2.0)
│   │   └── Logging/                ILog, FileLog, NullLog
│   ├── PyExcel.Bridge/           Frame transport (net48 + netstandard2.0)
│   │   ├── Framing.cs              Encode/decode wire frames
│   │   ├── FrameType.cs / Frame.cs / FramingExceptions.cs
│   │   ├── CanonicalJson.cs        Stdlib JSON encoder mirroring Python's json.dumps
│   │   ├── FrameTransport.cs       Stream/named-pipe framing wrapper
│   │   └── KernelSupervisor.cs     Spawn + HELLO/PING/SHUTDOWN over the pipe
│   ├── PyExcel.Kernel.Client/    Typed run/cancel/progress API over frames
│   ├── PyExcel.State/            Workbook registry, enabled/ribbon state, CustomXMLPart codec
│   ├── PyExcel.Excel/            Range ↔ Arrow marshalling, paste, import/export, charts
│   ├── PyExcel.Setup/            venv, pip, kernel extraction, project provisioning
│   ├── PyExcel.Forms/            WinForms dialogs (net48 + netstandard2.0 validation core)
│   ├── PyExcel.Ribbon/           ExcelRibbon subclass + Resources/RibbonUI.xml + customLogo.png (net48)
│   └── PyExcel.Addin/            .xll entry point — AddIn.cs + .dna (net48)
│
├── embedded/                     ── PYTHON KERNEL (shipped inside the .xll) ──
│   └── pyexcel/
│       ├── __init__.py
│       └── kernel/
│           ├── framing.py          Wire framing protocol (done + tested)
│           ├── transport.py        AF_UNIX (POSIX) / win32 (TODO) pipe client
│           ├── supervisor.py       HELLO + PING/PONG + SHUTDOWN event loop
│           └── __main__.py         `python -m pyexcel.kernel` entry point
│
├── tests/                        ── v2 TESTS (pytest + xUnit, cross-platform) ──
│   ├── conftest.py                 Puts embedded/ on sys.path
│   ├── kernel/                     Python tests (pytest)
│   │   ├── test_framing.py           Framing roundtrip + malformed-frame suite
│   │   └── test_cross_language_vectors.py   Golden hex vectors, paired with C#
│   └── PyExcel.Bridge.Tests/       C# tests (xUnit, net8.0)
│       ├── FramingTests.cs           Port of test_framing.py
│       ├── CrossLanguageVectorsTests.cs   Paired with test_cross_language_vectors.py
│       ├── FrameTransportTests.cs    MemoryStream + real named-pipe roundtrips
│       └── KernelSupervisorTests.cs  Integration: spawn python, HELLO/PING/SHUTDOWN
│
└── docs/
    └── v2-build.md               v2 build commands + Phase 1 exit gate
```

## History: the v1 VBA architecture (removed)

The original v1 add-in was VBA in `PyExcel.xlam`. `modRibbon.bas` handled ribbon
clicks → `python.bas` orchestrated a run → `xmlParsing.bas` serialized ranges →
a spawned `python.exe` ran `tools.py` → `pythonUtils.bas` + `chartBuilder.bas`
pasted results back, while VBA **polled `meta.xml`** on disk for run status.

That design's defining property was **file-based IPC + polling** — the source of
most v1 defects (locale-dependent number text, date-epoch guessing, polling
hangs, orphaned processes). The architecture below replaces it wholesale; the v1
VBA sources, `PyExcel.xlam`, and the on-disk Python runtime have been removed
from the repository.

## Architecture

**Process model.** Excel loads the `.xll` once (`AddIn.AutoOpen`). A single
long-lived **Python kernel** subprocess is spawned lazily and supervised; the
C# side talks to it over a **named pipe**.

**Projects** (see [`ROADMAP.md`](ROADMAP.md) for live build status):

| Project | Owns | Status |
| --- | --- | --- |
| `PyExcel.Common` | Logging (`ILog`), shared types | Phase 1 ✅ |
| `PyExcel.Ribbon` | Ribbon callbacks, `RibbonUI.xml` | Phase 1 skeleton ✅ |
| `PyExcel.Addin` | `.xll` entry point, service lifetime | Phase 1 skeleton ✅ |
| `PyExcel.Bridge` | Frame transport, named pipe, `KernelSupervisor` | Phase 2 — framing ✅, transport ✅, supervisor ✅ (POSIX; win32 client pending) |
| `PyExcel.Kernel.Client` | Typed run/cancel/progress API over frames | Phase 2 |
| `PyExcel.State` | Workbook registry, enabled state, ribbon state | Phase 3 |
| `PyExcel.Excel` | Range ↔ Arrow marshalling, paste, import/export, charts (`ChartBuilder`) | Phase 4–6 |
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
4. `PyExcel.Excel` writes the Arrow result tables to the mapped ranges; a
   Plotly figure return rides the wire as a JSON **chart spec**
   (`pyexcel.kernel.chart` → `ChartSpecParser` → `ChartBuilder` renders a
   native chart); a Matplotlib figure rides as a rendered SVG/PNG the host
   embeds as a picture.

### The kernel protocol

`embedded/pyexcel/kernel/framing.py` defines the wire format: a length-prefixed
binary frame = `body_len · type · meta_json · payloads`. Frame types
(`HELLO`, `PING`/`PONG`, `RUN_REQUEST`/`RUN_RESULT`, `PROGRESS`, `LOG`,
`CANCEL`, `ERROR`, …) are a stable enum. `PROTOCOL_VERSION = 2`. The C#
counterpart (`PyExcel.Bridge/Framing.cs`) **must match this byte-for-byte** —
build cross-language conformance tests. Frame sizes are bounded (256 MiB) on
both encode and decode to bound memory against a malformed peer.

## Where to make changes (navigation)

| To change… | Where |
| --- | --- |
| A ribbon button's behaviour | `PyExcel.Ribbon/PyExcelRibbon.cs` |
| Ribbon layout (tabs/groups) | `src/PyExcel.Ribbon/Resources/RibbonUI.xml` |
| How a run is orchestrated | `PyExcel.Kernel.Client` + `PyExcel.Bridge` |
| Range → Python data marshalling | `PyExcel.Excel` + `pyexcel.kernel` (Arrow) |
| Writing results back to ranges | `PyExcel.Excel` |
| Chart rendering | `embedded/pyexcel/kernel/chart.py` (figure → spec) + `PyExcel.Excel/ChartBuilder.cs` (spec → chart) |
| The `transform()` contract / runner | `embedded/pyexcel/kernel/worker.py` |
| The wire protocol | `embedded/pyexcel/kernel/framing.py` ↔ `PyExcel.Bridge/Framing.cs` |
| Import / Export / Paste | `PyExcel.Excel` services |
| First-run setup / venv | `PyExcel.Setup` |
| Dialogs | `PyExcel.Forms` |
| Workbook/ribbon state | `PyExcel.State` |
| Logging | `PyExcel.Common/Logging/FileLog.cs` |
| Project dependencies | `requirements.txt` (kernel set) |

## Build & test

Cross-platform (Linux/macOS — what CI runs):
```bash
dotnet build src/PyExcel.Common/PyExcel.Common.csproj --framework netstandard2.0 -c Release
dotnet build src/PyExcel.Bridge/PyExcel.Bridge.csproj  --framework netstandard2.0 -c Release
dotnet test  tests/PyExcel.Bridge.Tests/PyExcel.Bridge.Tests.csproj -c Release
pytest tests/
```
Windows (produces the `.xll`):
```powershell
dotnet restore PyExcel.sln
dotnet build PyExcel.sln -c Release --no-restore
```
Full build prerequisites and the Phase 1 exit gate are in [`docs/v2-build.md`](docs/v2-build.md).
The CI workflow is `.github/workflows/ci.yml` — Linux runs the cross-platform commands above and Windows builds the full solution.

## Conventions

- **Logging.** Writes to `%TEMP%\PyExcel_Debug.log`, format
  `[YYYY-MM-DD hh:mm:ss.fff] [LEVEL] message`, via `ILog` (`PyExcel.Common`).
- **Versioning.** .NET assemblies: SemVer `2.0.0-alpha` (`Directory.Build.props`).
  Python kernel: PEP 440 `2.0.0a0` (`embedded/pyexcel/__init__.py`) — the same
  release in each ecosystem's convention.
- **Phase model.** Phases 1–9 are defined in [`ROADMAP.md`](ROADMAP.md) — the
  single source of truth. `// PHASE n` comments in C# must match it.
- **SAFE-1.** Ribbon callbacks must never block on the kernel pipe — enqueue and return.
- **Branching.** Feature work happens on a dedicated branch, not `main`.

## Known issues & tech debt

The full audit — bugs, dead code, inefficiencies, UX gaps — is folded into
[`ROADMAP.md`](ROADMAP.md) under *Bug disposition* and *Cut list*, classified by
whether the rewrite's architecture deletes them or whether they must be fixed
during the port.
