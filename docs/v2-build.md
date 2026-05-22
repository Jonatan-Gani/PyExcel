# PyExcel v2 — Build & Project Layout

This document describes how to build PyExcel v2 from source. v2 is the
.NET / Excel-DNA rewrite of the v1 `.xlam`. The canonical codebase map is
[`ARCHITECTURE.md`](../ARCHITECTURE.md); the phased build plan and task list
is [`ROADMAP.md`](../ROADMAP.md).

## Solution layout

```
PyExcel/
├── PyExcel.sln                       Top-level solution
├── Directory.Build.props             Shared MSBuild settings (versioning, lang)
├── src/
│   ├── PyExcel.Common/               net48 + netstandard2.0 — logging, types
│   ├── PyExcel.Ribbon/               net48 — ExcelRibbon subclass, RibbonUI.xml
│   └── PyExcel.Addin/                net48 — .xll entry point + .dna file
├── embedded/
│   └── pyexcel/                      Python kernel sources (shipped inside .xll)
│       └── kernel/                   Framing, supervisor, worker (in progress)
├── tests/
│   └── kernel/                       pytest, cross-platform
├── docs/
│   └── v2-build.md                   This file
└── build/                            Sign + pack scripts (planned — not yet created)
```

Phases 2–9 build out `PyExcel.Bridge`, `PyExcel.Kernel.Client`, `PyExcel.State`,
`PyExcel.Excel`, `PyExcel.ChartBuilder`, `PyExcel.Setup`, `PyExcel.Forms`, and the
Python kernel, then retire v1. **[`ROADMAP.md`](../ROADMAP.md) is the single
source of truth for phase numbering, scope, and task status** — the `// PHASE n`
comments in the C# must be kept in sync with it.

## Prerequisites (Windows, for producing the `.xll`)

- Visual Studio 2022 with the .NET desktop workload, **or** the standalone
  Build Tools for Visual Studio with the same workload.
- .NET Framework 4.8 targeting pack (ships with VS).
- .NET 6+ SDK (any version; used to drive `dotnet build`).
- Excel for Windows 2016 / 2019 / 365 (x64) for end-to-end testing.

## Prerequisites (Linux/macOS, for CI builds of the cross-platform projects)

- .NET 8 SDK.
- Python 3.10+.
- `pytest`.

You cannot produce the final `.xll` on Linux — `PyExcel.Addin` and
`PyExcel.Ribbon` target `net48` and depend on packages that resolve only
under Windows. The cross-platform CI build covers `PyExcel.Common`
(`netstandard2.0` slice) plus the entire Python kernel test suite.

## Build commands

### Cross-platform (Linux/macOS, what runs in CI)

```bash
# Verify the netstandard2.0 slice of PyExcel.Common compiles.
dotnet build src/PyExcel.Common/PyExcel.Common.csproj \
    --framework netstandard2.0 \
    --configuration Release

# Run the Python kernel test suite.
pytest tests/
```

### Windows — Phase 1 exit deliverable

```powershell
# Build everything; ExcelDnaPack runs as a post-build target on PyExcel.Addin
# and produces PyExcel-AddIn64.xll under src/PyExcel.Addin/bin/Release/.
dotnet restore PyExcel.sln
dotnet build PyExcel.sln --configuration Release --no-restore
```

Phase 1 exit gate (manual):

1. Open Excel (any blank workbook).
2. File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse →
   select `src\PyExcel.Addin\bin\Release\net48\PyExcel-AddIn64.xll`.
3. A **Python** ribbon tab appears with the five groups: Main, Python,
   Import, Export, Paste.
4. Click **Read Me** in the Main group — an `xlcAlert` dialog appears
   identifying the build as v2 alpha.
5. Every other button is disabled or stubbed (logs to
   `%TEMP%\PyExcel_Debug.log` when clicked).

If `PyExcelEnabled` is checked anywhere in Excel state from a v1 workbook,
that workbook will fail the v2 `RibbonEnabled` guard — Phase 1 doesn't yet
include the v1→v2 migration (Phase 9 work).

## Project targeting summary

| Project | TargetFrameworks | Builds on Linux? | Notes |
|---|---|---|---|
| `PyExcel.Common` | `net48;netstandard2.0` | Yes (`netstandard2.0` slice) | Pure types + logger |
| `PyExcel.Ribbon` | `net48` | No | Needs ExcelDna.Integration |
| `PyExcel.Addin` | `net48` | No | Produces the `.xll`; needs ExcelDna.AddIn |

## Logs

When the `.xll` is loaded, it writes to `%TEMP%\PyExcel_Debug.log`. v1
wrote to the same path, so post-mortems can mix logs from both versions
during the migration window. Format: `[YYYY-MM-DD hh:mm:ss.fff] [LEVEL] message`.
