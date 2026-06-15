# Logging

## Macro
A minimal, cross-platform logging seam used everywhere in the add-in: one
interface and two implementations (file-backed and no-op).

## Files
### ILog.cs
The logging interface: five severity methods (Trace, Debug, Info, Warn, Error) taking a
`string.Format`-style message and optional exception. Inputs: a message and optional args
/ exception. Output: none (logging side effect).

### FileLog.cs
Append-only text-file logger writing one timestamped line per event to
`%TEMP%\PyExcel_Debug.log` (or an override path), using a short-lived writer per call and
swallowing all failures. Inputs: an optional log-file path. Output: appended UTF-8 lines.

### NullLog.cs
No-op singleton logger for tests and headless runs. Inputs: ignored. Output: none.

## Subdirectories
None.
