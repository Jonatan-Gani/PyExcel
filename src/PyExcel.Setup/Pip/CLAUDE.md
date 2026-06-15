# Pip

## Macro
Invoking pip inside the project venv and verifying the kernel's dependencies are
importable.

## Files
### PipRunner.cs
Invokes pip inside the project venv via `python -m pip`: Install (from a requirements
file), UpgradePip (best-effort), and Show (query a package). Inputs: the venv python
executable, a requirements file path (Install), or a package name (Show). Output: a
`ProcessRunResult` (exit code, stdout, stderr).

### DependencyVerifier.cs
Verifies the kernel's required modules (pandas, numpy, pyarrow, plotly, matplotlib) import
by running `python -c "import X"` for each; strict (all must import). Inputs: the venv
python executable and an optional per-module timeout. Output: a
`DependencyVerificationResult` (all-importable flag + per-module status).

## Subdirectories
None.
