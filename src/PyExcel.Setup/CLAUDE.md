# PyExcel.Setup

## Macro
Headless provisioning of a workbook's project environment: probe system Python,
create a project-local venv, extract the embedded kernel package, upgrade pip,
install requirements, and verify the imports. The top-level files orchestrate and
report; the subpackages own each discrete stage. Cross-platform and process-
driven (no COM); the WinForms wizard that hosts it lives in `PyExcel.Forms`.

## Files
### SetupService.cs
The top-level facade running the stages in order (resolve+classify path, ensure project
dir, probe Python, provision venv, extract kernel, upgrade pip, install requirements,
verify deps); a failed stage short-circuits the rest. Inputs: a project path string.
Output: a structured `SetupResult` (per-stage outcomes + success flag); never throws.

### SetupReport.cs
Formats a `SetupResult` into a human-readable summary — a one-line headline plus a
per-stage `[ok]`/`[fail]` transcript. Inputs: a `SetupResult`. Output: a formatted string.

### ProjectScaffolder.cs
Prepares user-facing project directories: creates `userScripts/`, seeds an `example.py`
on first run, and writes a project README; idempotent (never overwrites). Inputs: a
project directory path. Output: the `userScripts` path; writes `example.py` and
`README.md`.

## Subdirectories
- **Python/** (own CLAUDE.md) — system-Python probe and venv provisioner.
- **Kernel/** (own CLAUDE.md) — extracts the embedded kernel package from assembly
  resources to disk.
- **Pip/** (own CLAUDE.md) — pip invocation and kernel dependency verification.
- **Paths/** (own CLAUDE.md) — project path normalisation/classification (Local / UNC /
  OneDrive).
