# Shell

## Macro
Process and shell execution shared by Setup and the ribbon: a cross-platform
child-process runner and a Windows-only shell-verb launcher.

## Files
### ProcessRunner.cs
Cross-platform child-process executor that streams stdout/stderr to an `ILog` while
accumulating the full streams for the caller (Windows `CommandLineToArgvW` quoting for
paths with spaces). Inputs: an executable path, an argument list, an optional working
directory, environment overlay, timeout, and cancellation token. Output: a
`ProcessRunResult` (exit code, captured stdout, captured stderr).

### ShellLauncher.cs
Windows-only thin wrapper over shell verbs — open a file with its registered handler,
reveal a file in Explorer, open a folder (net48, `UseShellExecute`). Inputs: file or
folder paths (validated for null/empty, not existence). Output: none (launches external
programs).

## Subdirectories
None.
