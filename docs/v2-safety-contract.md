# PyExcel v2 — Safety Contract

This document is the binding safety specification for the v2 Python↔Excel bridge.
Every invariant below is **testable**, **enforced at a named call site**, and has a
defined **failure mode** that never crashes Excel.

The contract exists because Excel's event surface is reactive: tab switches, window
focus changes, the 10 s watchdog, and ribbon `getEnabled`/`getText`/`getLabel`
callbacks all fan out frequently and on the UI thread. Without explicit guards,
attaching a long-running Python kernel to that surface produces hangs, double
spawns, and silent state corruption. The rules below exist so that doesn't happen.

---

## 1. Event surface we live inside

Verified from `src/module/HostManager.bas` and `src/module/CAppEvents.cls`. Any
v2 code that runs on the UI thread MUST assume each of these fires unpredictably:

| Event | Source | Frequency | What it does today |
|-------|--------|-----------|--------------------|
| `App_WorkbookActivate` | `CAppEvents.cls:54` | every workbook switch | `HostManager_ActivateWorkbook` → ribbon refresh; if `PyExcelEnabled=1` schedules `VerifyProjectVersion` via `Application.OnTime + 1s` |
| `App_SheetActivate` | `CAppEvents.cls:61` | every tab click | `HostManager_ActivateSheet` → ribbon refresh |
| `App_WindowActivate` | `CAppEvents.cls:69` | every Alt-Tab back into Excel | `HostManager_RefreshRibbonOnly` (no version check) |
| `App_WorkbookBeforeClose` | `CAppEvents.cls:77` | workbook close | `HostManager_UnregisterHost` |
| `HostManager_Watchdog` | `HostManager.bas:355` | every 10 s via `Application.OnTime` | Self-heals event sink + registries; refreshes ribbon if `RibbonIsEnabled` |
| Ribbon `getEnabled` / `getText` / `getLabel` callbacks | invoked by `Ribbon.InvalidateControl` inside `HostManager_RibbonRefreshAll` (`HostManager.bas:272-321`) | up to 20+ per refresh, synchronously on UI thread | reads workbook state, returns immediately |

The watchdog and `HostManager_RibbonRefreshAll` are the highest-frequency callers.
Anything they touch must be O(1), allocation-light, and pipe-free.

---

## 2. Invariants

Each invariant has an ID. Code that enforces it cites the ID in a comment.
Each invariant has a defined enforcement site and a defined failure mode.

### SAFE-1. Ribbon callbacks never touch the pipe

**Statement.** No `getEnabled`, `getText`, `getLabel`, `getItemCount`, `getItemLabel`,
or `onChange` ribbon callback ever calls `kernelClient.*` or performs pipe I/O,
directly or transitively.

**Enforcement.** `kernelClient.bas` exposes only a single public entry point,
`KernelClient_RunJob`, which is callable solely from `OnRunPython`
(the `btnRun` action) and the equivalent Action-replay sub. All other
ribbon callbacks read VBA-side cached state populated outside callbacks.

**Failure mode.** If a callback ever calls `KernelClient_*` while a job is
in flight, the gate `KernelClient_AssertNotInRibbonCallback` raises and
the call is logged. Excel keeps running; ribbon refresh completes.

**Test.** `tests/vba/safe_1_ribbon_freeze.bas` drives 1000 simulated ribbon
refreshes during a 30 s job and asserts the job completes within +5 % of
its no-load duration and the UI thread is never blocked > 100 ms.

### SAFE-2. Kernel spawn is lazy, never on workbook/sheet activate

**Statement.** A Python kernel process is started only by the first
`KernelClient_RunJob` call after Excel start (or after a kernel death).
`Workbook_Activate`, `Sheet_Activate`, `Window_Activate`, and the watchdog
never spawn a kernel.

**Enforcement.** `kernelLifecycle.EnsureKernel` is called only from
`KernelClient_RunJob`. `HostManager_ActivateWorkbook`,
`HostManager_ActivateSheet`, `HostManager_RefreshRibbonOnly`, and
`HostManager_Watchdog` MUST NOT reference `kernelLifecycle`.

**Failure mode.** If `EnsureKernel` is ever called from a non-Run context,
the function detects the call stack (sentinel global `gKC_InRunContext`)
and returns `Nothing` with a log line. The caller treats this as "kernel
not available" and falls back to v1 spawn path.

**Test.** Static grep gate in CI: `grep -E '\bkernelLifecycle_|\bKernelClient_' src/module/*.bas`
must produce zero hits outside `kernelClient.bas`, `kernelLifecycle.bas`,
`python.bas` (in the `OnRunPython` chain), and `Update.bas` (drain only).

### SAFE-3. EnsureKernel is idempotent and serialized per Excel.Application

**Statement.** Multiple concurrent calls to `EnsureKernel` from the same
Excel process produce exactly one live kernel. Calls during a spawn-in-flight
block on that spawn, not start a new one.

**Enforcement.** A module-level `gKL_State As LONG` (0=Idle, 1=Spawning,
2=Ready, 3=Draining) with `InterlockedCompareExchange` semantics implemented
via a wrapper around `Application.OnTime` continuations. Spawn is owned by
the goroutine-equivalent that wins the CAS from Idle→Spawning.

**Failure mode.** If state machine wedges (spawn never completes, never
errors), the bounded total spawn timeout of 8 s triggers a forced reset to
Idle and an error toast. No deadlock possible.

**Test.** `tests/kernel/test_lifecycle.py` simulates 10 concurrent
EnsureKernel calls via threads against a mock transport and asserts exactly
one spawn() invocation.

### SAFE-4. All pipe waits are message-pumped, bounded, and cancellable

**Statement.** Any VBA code that waits for a pipe event pumps the Windows
message queue (so Excel UI stays responsive) and has a hard wall-clock
ceiling. No `Sleep`-only loops.

**Enforcement.** `kernelClient.WaitFrame(timeoutMs)` uses
`MsgWaitForMultipleObjectsEx(1, hEvent, timeoutMs, QS_ALLINPUT, MWMO_ALERTABLE)`
with `DoEvents` between iterations and a 100 ms tick ceiling per inner wait
even when caller passes a longer timeout. The user-visible ufProgress dialog
exposes a Cancel button that signals a per-job `CancelEvent` HANDLE the wait
also subscribes to.

**Failure mode.** On timeout: caller receives a `KernelError_Timeout` and
the worker is sent a `Cancel` frame; if no `Bye` ack within 1 s, supervisor
terminates the worker.

**Test.** `tests/kernel/test_blocking.py` runs a fake worker that holds
output for N seconds while the harness asserts wall-clock liveness on the
"VBA side" mock (no event-loop starvation).

### SAFE-5. Reentrancy of `Run` is forbidden at the workbook level

**Statement.** `btnRun_Click` short-circuits with a toast if a job is
already in flight for the same workbook. A different workbook may run
concurrently.

**Enforcement.** Per-workbook lock stored on the workbook as a custom
document property `PyExcel_RunInFlight = "<run_id>"`. `OnRunPython`'s
first action is `If Len(GetWorkbookValue(wb, "PyExcel_RunInFlight")) > 0 Then
ShowBusyToast: Exit Sub`. Cleared by `KernelClient_RunJob` on completion or
error.

**Failure mode.** If a previous run died without clearing the flag and
the kernel reports no such `run_id`, the flag is auto-released on the
next button press with a `[recovered stale lock]` log line.

**Test.** `tests/vba/safe_5_reentrancy.bas` queues two `btnRun` events
5 ms apart via `Application.OnTime`; the second must produce the busy
toast, not a second spawn.

### SAFE-6. WorkbookBeforeClose and Application.Quit always complete in ≤ 3 s

**Statement.** Closing a workbook or quitting Excel never hangs on
v2 code, regardless of kernel state.

**Enforcement.** `App_WorkbookBeforeClose` calls `KernelLifecycle_Drain(wb,
maxMs:=3000)` which (a) writes a `Cancel` frame to any in-flight job for
this workbook, (b) waits up to 2 s with message pumping for `Bye` acks,
(c) `TerminateProcess` on any unresponsive worker, (d) `CloseHandle` on
pipes, (e) returns. Total wall-clock cap: 3 s. `App_Quit` calls
`KernelLifecycle_DrainAll(maxMs:=3000)` with the same semantics.

**Failure mode.** None observable by user. Worker zombies are not possible
because the kernel parent-PID watchdog (see SAFE-7) self-terminates.

**Test.** Manual: kill Excel via Task Manager mid-job, observe no
`pyexcel_kernel` process in tasklist after 10 s.

### SAFE-7. Kernel detects orphaning and self-terminates within 2 s

**Statement.** If the Excel process that spawned the kernel goes away
(crash, force-kill, RDP disconnect cleanup), the kernel exits within 2 s
without writing to disk.

**Enforcement.** `pyexcel.kernel.supervisor` runs a 1 s ticker calling
`os.kill(parent_pid, 0)`. On `ProcessLookupError` (POSIX) or
`OSError(ERROR_INVALID_PARAMETER)` (Windows on dead PID), supervisor
sends SIGTERM to its workers, joins for 500 ms, then `os._exit(0)`.
Workers do not write archive/temp files after receiving the supervisor
shutdown signal.

**Failure mode.** Worst case: kernel hangs in a non-Python C call. Mitigated
by the kernel being a separate process — Excel close is not blocked, and
the OS reaps the orphan when its session ends.

**Test.** `tests/kernel/test_orphan.py` spawns a fake parent + a real
kernel, kills the parent with SIGKILL, asserts the kernel exits within
2 s.

### SAFE-8. Pipe ACL restricts the kernel to the current user's SID

**Statement.** No other local user account, including SYSTEM-context
services running on the same machine, can connect to a PyExcel kernel pipe.

**Enforcement.** `kernel.transport.NamedPipeTransport._make_security_attributes`
constructs an explicit DACL with two ACEs: `GENERIC_READ | GENERIC_WRITE`
for the current user SID (obtained via `GetTokenInformation(TokenUser)`),
and `OWNER_RIGHTS` for the same SID. No `EVERYONE`, no `Authenticated Users`,
no `INTERACTIVE`. On accept, `GetNamedPipeClientProcessId` →
`OpenProcessToken` → `GetTokenInformation(TokenUser)` → SID match required.

**Failure mode.** Non-matching client: `DisconnectNamedPipe` and log entry.
Pipe creation failure (rare, e.g. extreme name collision): kernel logs and
exits with status `PIPE_ACL_FAILURE`; `kernelClient` reports kernel-start
failure to user, v1 fallback can be enabled per workbook.

**Test.** `tests/kernel/test_pipe_acl.py` runs only on Windows CI; on Linux
the transport abstraction substitutes a Unix-socket transport with
`0600` permissions and the equivalent assertion.

### SAFE-9. Update flow drains kernels before pip operations

**Statement.** A `pip install` / `pip uninstall` triggered by `Update.bas`
never races with a running kernel that has mapped `.pyd` files.

**Enforcement.** Before any pip call, `Update.bas` calls
`KernelLifecycle_DrainAll(maxMs:=5000)` (same drain protocol as
WorkbookBeforeClose but with a longer ceiling). Pip is invoked only after
drain returns. Re-spawn is deferred until next Run.

**Failure mode.** If a worker refuses to drain in 5 s, `TerminateProcess`
runs. Worst case for the user: any in-flight Run loses its result; the
Update proceeds. The user is informed via a confirmation dialog before
the update begins.

**Test.** `tests/integration/test_update_drain.py` starts a long-running
fake job, triggers a simulated update, asserts the kernel exits before pip
runs.

### SAFE-10. State machine is explicit and observable

**Statement.** At any moment the kernel client knows which of these states
it is in: `Cold` (no kernel), `Spawning`, `Ready` (kernel alive, no job),
`Running` (job in flight), `Draining`, `Dead`. Transitions are logged.

**Enforcement.** `kernelLifecycle.bas` owns the state machine. All
transitions go through `KernelLifecycle_Transition(newState, reason)`,
which logs and updates the `gKL_State` global. `KernelClient_DebugDumpState`
is a public sub that prints the current state to the debug log on demand —
makes incident triage trivial.

**Failure mode.** Illegal transition (e.g., Running → Spawning) raises and
is logged with the offending transition. State is reset to `Cold` and a
toast tells the user to retry.

**Test.** `tests/kernel/test_state_machine.py` exhaustively walks legal
and illegal transitions; illegal ones must raise without partial state
mutation.

### SAFE-11. Kernel failure is degraded, not fatal

**Statement.** If kernel spawn fails (venv corrupted, python missing,
permission denied), Excel remains fully functional. The user sees a clear
error; non-Python ribbon features (Import, Export, Paste of pre-existing
artifacts) keep working.

**Enforcement.** `KernelClient_RunJob` wraps `EnsureKernel` in an error
handler. On failure, sets `gKC_KernelUnavailable = True`, surfaces a toast
with the diagnostic, and the Run button's `getEnabled` returns False until
the user clicks **Enable PyExcel** to re-run setup. The unavailable state
auto-clears on next workbook open.

**Failure mode.** None — this *is* the failure mode for upstream errors.

**Test.** `tests/integration/test_kernel_failure.py` corrupts the venv
shebang, runs the harness, asserts the workbook remains responsive and
the error surfaces visibly.

### SAFE-12. No global mutable state leaks across runs

**Statement.** A user script's module-level state (caches, imports done by
that user script) is reset between runs. Two consecutive runs of the same
script must produce identical output given identical input.

**Enforcement.** Workers import user scripts via
`importlib.util.spec_from_file_location` into a fresh module per run, then
discard. The kernel's own imports (pandas, plotly, pyarrow) are preserved
across runs — that's the entire point of the persistent kernel — but user
namespace is not.

**Failure mode.** A user script that *relies* on module-level caching
across runs (legitimate use case, rare) opts in via `@job(cache=True)`. The
default is reset.

**Test.** `tests/kernel/test_run_isolation.py` runs a script twice; the
second run must not see a global set by the first.

---

## 3. What this contract does NOT cover

- **Excel formula recalculation correctness** — out of scope; the bridge is
  Run-button driven, not cell-driven.
- **Cross-platform** — Windows only. Linux dev/test uses Unix sockets via the
  transport abstraction; Mac is not a target.
- **Sandboxing user scripts** — same trust model as v1. The contract does
  not protect against malicious user scripts run in the worker.
- **Network exposure** — the kernel listens on a named pipe only; no TCP in
  production. `PYEXCEL_DIAG_TCP=1` opens a localhost-bound port for
  diagnostics; it is a dev tool, not a deployment mode.
- **Multi-`Excel.exe` kernel sharing** — each Excel.exe gets its own kernel.

---

## 4. CI gates

Every PR that touches `src/module/kernel*.bas`, `src/embedded/pyexcel/kernel/`,
or `src/embedded/pyexcel/runtime/` must pass:

1. **Static safety lint** — `tests/ci/safety_lint.py`:
   - SAFE-2 grep gate (no `kernelLifecycle_` references in `HostManager.bas`
     or `CAppEvents.cls`).
   - SAFE-1 grep gate (no `KernelClient_` references in ribbon callback
     handlers — identified by signature `Sub <Name>(control As IRibbonControl, ...)`).
2. **Python unit tests** — `pytest tests/kernel tests/runtime`. Must pass on
   Linux (this is what we verify in CI) and on Windows (verified locally
   on the maintainer's box).
3. **Wire-protocol roundtrip** — fuzz test that frames-encoded then
   frames-decoded payloads of random shape/size are bit-identical.

Manual Windows/Excel integration tests are catalogued in
`tests/integration/README.md` and run by the maintainer before each tagged
release.

---

## 5. Versioning of this contract

Contract version is committed at the top of `src/embedded/pyexcel/kernel/framing.py`
as `PROTOCOL_VERSION` and at the top of `src/module/kernelClient.bas` as
`KC_PROTOCOL_VERSION`. `Hello` frames carry the version both sides
advertise. Mismatch → kernel client refuses to attach and surfaces a
diagnostic toast. Bumping any invariant or wire-format detail bumps both
constants in the same PR.
