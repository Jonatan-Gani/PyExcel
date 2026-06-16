# Python

## Macro
Locating a usable system Python and creating the per-project virtual environment
Setup installs into.

## Files
### SystemPythonProbe.cs
Locates a usable system Python from a `PYEXCEL_PYTHON` override or a PATH search,
rejecting the Windows Store stub by path prefix and file size and verifying with
`python --version`. Inputs: an optional environment-variable override. Output: a
`PythonProbeResult` (found flag, executable path, version banner, or failure reason).

### VenvProvisioner.cs
Creates a per-project venv at `.pyexcel-venv` via `python -m venv`; idempotent, and
recreates a corrupted venv (directory present but no executable). Inputs: a project
directory and the system Python executable. Output: a `VenvProvisionResult` (venv
directory, venv python path, outcome).

## Subdirectories
None.
