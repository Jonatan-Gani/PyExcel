# PyExcel.Common

## Macro
Utilities shared across the add-in, setup, and ribbon: project-directory
resolution plus the `Shell/` and `Logging/` subpackages. Cross-platform where it
can be; the Windows-only shell launcher is isolated in `Shell/`.

## Files
### ProjectDirectory.cs
Resolves the local directory where PyExcel stores a workbook's per-project Python
environment (`.pyexcel-venv`, `.pyexcel-kernel`). Inputs: a workbook folder path or URL
from Excel; honours a `PYEXCEL_PROJECT_DIR` override and maps cloud/URL workbooks to a
`%LOCALAPPDATA%\PyExcel` fallback. A two-arg overload `Resolve(storedProjectDir,
workbookDir, directoryExists?)` adds self-healing: it honours a stored project folder only
while it still exists, else falls back to the workbook's own folder (so a moved project
folder re-resolves). Output: a normalised absolute path, or null for an unsaved workbook.

## Subdirectories
- **Shell/** (own CLAUDE.md) — cross-platform child-process runner and the Windows shell
  launcher. Boundary: takes an executable + args, returns captured stdout/stderr/exit.
- **Logging/** (own CLAUDE.md) — the `ILog` interface and its file/null implementations.
  Boundary: takes severity + message, writes log lines.
