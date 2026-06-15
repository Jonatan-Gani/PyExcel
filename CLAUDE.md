# PyExcel

## Macro
PyExcel is a Windows Excel add-in (v2, built on .NET / Excel-DNA) that runs
user-authored Python `transform(inputs)` functions against workbook ranges in an
isolated, project-local Python environment. The user selects input ranges, picks
a script, and clicks **Run**; the add-in marshals the ranges to a supervised
Python kernel subprocess, executes the script, and writes the returned tables,
lists, scalars, and charts back into Excel. Two planes cross the boundary: a
control plane of length-prefixed binary frames with canonical-JSON metadata, and
a data plane of Arrow IPC streams carrying shape metadata so results spill back
into the right cell geometry. Runs are non-reactive — nothing executes until Run
is clicked. The repo also holds the embedded Python kernel package shipped inside
the `.xll`, plus Import/Export/Paste data utilities and per-workbook state that
persists inside the workbook's CustomXMLPart.

The codebase multi-targets `net48` (the in-proc `.xll`, COM + WinForms, gated
behind `#if NETFRAMEWORK`) and `netstandard2.0` (pure libraries that build and
test on Linux CI). The kernel wire format is mirrored byte-for-byte between
`src/PyExcel.Bridge/Framing.cs` (C#) and `embedded/pyexcel/kernel/framing.py`
(Python); changes to one must track the other.

## Project tree
```
PyExcel/
├── ARCHITECTURE.md, README.md, ROADMAP.md      project docs (human-authored)
├── PyExcel.sln, Directory.Build.props          solution + shared MSBuild settings
├── requirements.txt                            kernel Python dependencies
├── .github/workflows/ci.yml                    CI (Linux cross-platform slice)
├── docs/                                        build + phase design docs
│   ├── v2-build.md, phase3-and-4-completion.md, phase4-handoff.md
├── embedded/pyexcel/                            Python kernel package (shipped in the .xll)
│   ├── __init__.py
│   └── kernel/   [CLAUDE.md]                    framing, transport, arrow_io, chart, worker, supervisor
├── src/
│   ├── PyExcel.Common/   [CLAUDE.md]            ProjectDirectory; Shell/ + Logging/ subdirs
│   │   ├── Shell/   [CLAUDE.md]                 ProcessRunner, ShellLauncher
│   │   └── Logging/ [CLAUDE.md]                 ILog, FileLog, NullLog
│   ├── PyExcel.Bridge/   [CLAUDE.md]            framing protocol + KernelSupervisor (pipe + subprocess)
│   ├── PyExcel.Kernel.Client/ [CLAUDE.md]       typed RPC client over the supervisor
│   ├── PyExcel.Excel/    [CLAUDE.md]            marshalling, run drivers, CSV/charts, KernelHost
│   ├── PyExcel.State/    [CLAUDE.md]            per-workbook state, persistence codecs, run archive
│   ├── PyExcel.Forms/    [CLAUDE.md]            WinForms dialogs + cross-platform validators
│   ├── PyExcel.Ribbon/   [CLAUDE.md]            Excel-DNA ribbon callbacks (Resources/RibbonUI.xml)
│   ├── PyExcel.Addin/    [CLAUDE.md]            add-in entry, COM event sink, workbook persistence
│   └── PyExcel.Setup/    [CLAUDE.md]            venv/kernel/pip provisioning; Python/ Kernel/ Pip/ Paths/ subdirs
└── tests/
    ├── conftest.py                              adds embedded/ to sys.path
    ├── kernel/   [CLAUDE.md]                    pytest suite for the Python kernel
    └── PyExcel.Bridge.Tests/ [CLAUDE.md]        xUnit suite for the cross-platform C# slice
```

## Root files
### PyExcel.sln
Visual Studio solution tying together every `src/PyExcel.*` project and the C#
test project. Inputs/Output: build metadata only.

### Directory.Build.props
Shared MSBuild settings applied to every project: `LangVersion 10`, nullable
enabled, warnings-as-errors, version stamps, deterministic build. Inputs/Output:
build configuration only — note that warnings-as-errors means any nullable or
unused-symbol warning fails the build.

### requirements.txt
Python dependencies installed into each project venv (pandas, numpy, pyarrow,
plotly, matplotlib). Inputs/Output: consumed by `PyExcel.Setup` at provision time.

### ARCHITECTURE.md / README.md / ROADMAP.md
Human-authored prose: codebase map, user-facing authoring/usage guide, and live
status/roadmap. Not part of the machine-owned context map.

### .github/workflows/ci.yml
CI definition running the Linux cross-platform slice (`dotnet build`/`dotnet test`
of the netstandard projects + `pytest tests/`). Inputs/Output: CI orchestration.

## Subdirectories
- **embedded/** — the Python kernel package shipped inside the `.xll` and
  extracted to each project environment on Setup. Its real logic lives in
  `embedded/pyexcel/kernel/` (own CLAUDE.md); `embedded/pyexcel/__init__.py` is
  just the package version banner. Boundary: spawned as `python -m pyexcel.kernel`,
  connects back over a named pipe, exchanges frames + Arrow payloads.
- **src/** — all C# projects (see tree for the per-project one-liners; each has
  its own CLAUDE.md). Boundary inputs: Excel COM (ranges, workbook events) and
  the kernel pipe; outputs: cells written, charts built, state persisted in the
  workbook.
- **tests/** — `tests/kernel/` is the pytest suite for the Python kernel;
  `tests/PyExcel.Bridge.Tests/` is the xUnit suite covering the cross-platform C#
  slice (codecs, validators, planners, framing, Arrow, kernel client/supervisor).
  Both run on Linux CI. `conftest.py` puts `embedded/` on `sys.path`.
- **docs/** — build instructions and phase design/handoff notes (human-authored,
  no proprietary scripts).

## Conventions
- Write correctly-typed, efficient code.
- Prefer vectors and matrices over loops.
- Input/output sections describe data, never code.
- Windows-only surface (COM, WinForms, Excel-DNA) sits behind `#if NETFRAMEWORK`;
  everything testable is kept cross-platform so it runs on Linux CI.
- Warnings are errors with nullable enabled — new C# must be warning-clean.
- The kernel wire format is mirrored in `Framing.cs` and `framing.py`; keep them
  byte-for-byte compatible.

## Recommended skills
- **context-map-builder** — used to generate these CLAUDE.md files; re-run it
  after structural changes (new projects, moved/renamed proprietary scripts) so
  the map stays current.
- No other skill usage detected in this project.

## References
- Task list: TASKS.md
- Preserved human notes: projectNotes.md (per directory)
