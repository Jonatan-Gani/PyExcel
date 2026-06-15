using System;
using System.Threading;
using PyExcel.Bridge;
using PyExcel.Kernel.Client;
using PyExcel.State;

namespace PyExcel.Excel;

/// <summary>
/// Process-wide lifecycle wrapper around a single <see cref="KernelSupervisor"/>
/// + <see cref="KernelClient"/> pair. The first <see cref="Client"/> access
/// boots the kernel; subsequent calls reuse it. Disposal tears the
/// supervisor down — call it from the add-in unload hook.
///
/// <para>This is a stop-gap for Phase 4. Phase 3 will move ownership into
/// per-workbook <c>StateService</c> so each workbook has its own kernel
/// (and switching workbooks doesn't cross-contaminate Python module
/// caches). Keep the surface narrow so that migration is mechanical.</para>
///
/// <para>The boot path itself is thread-safe via <see cref="Lazy{T}"/>'s
/// <see cref="LazyThreadSafetyMode.ExecutionAndPublication"/> default —
/// concurrent first-access from multiple Excel calc threads will produce
/// one kernel, not many. Subsequent <c>Client.Run</c> calls serialise
/// inside <see cref="KernelSupervisor"/>'s exchange semaphore.</para>
/// </summary>
public sealed class KernelHost : IDisposable
{
    private static readonly Lazy<KernelHost> s_default = new(() => new KernelHost());

    /// <summary>Process-wide default host. Suitable for Phase 4; Phase 3
    /// callers should construct their own instances instead.</summary>
    public static KernelHost Default => s_default.Value;

    // _booted is swapped out by Restart()/GetBooted when the kernel dies, so it
    // is mutable and every access goes through _gate. _gate also serialises the
    // boot itself (as the old Lazy default did) and the dead-kernel reboot, so
    // two callers can't race two kernels into existence.
    private readonly object _gate = new();
    private Lazy<Booted> _booted;
    private int _disposed;

    public KernelHost()
    {
        _booted = new Lazy<Booted>(Boot, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>The typed client. First access boots the kernel; a dead kernel
    /// is transparently replaced with a fresh one.</summary>
    public KernelClient Client => GetBooted().Client;

    /// <summary>The supervisor (process + pipe owner).</summary>
    public KernelSupervisor Supervisor => GetBooted().Supervisor;

    /// <summary>True once the kernel has been booted at least once.</summary>
    public bool IsStarted
    {
        get { lock (_gate) { return _booted.IsValueCreated; } }
    }

    /// <summary>
    /// Dispose the current kernel (if booted) and arm a fresh boot on next use.
    /// Use after the kernel is known unusable — e.g. the hard-cancel escalation
    /// in <c>PyExcel.Kernel.Client.KernelClient</c> killed a wedged child that
    /// couldn't honour a Cancel. Idempotent; a no-op when never booted or
    /// already disposed. The next <see cref="Client"/>/<see cref="Supervisor"/>
    /// access boots the replacement.
    /// </summary>
    public void Restart()
    {
        if (Volatile.Read(ref _disposed) != 0) return;
        lock (_gate)
        {
            if (_booted.IsValueCreated)
            {
                try { _booted.Value.Supervisor.Dispose(); } catch { /* best-effort */ }
            }
            _booted = new Lazy<Booted>(Boot, LazyThreadSafetyMode.ExecutionAndPublication);
        }
    }

    /// <summary>
    /// Idempotent shutdown. Safe to call multiple times; safe to call
    /// before the kernel was ever booted (no-op in that case).
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        lock (_gate)
        {
            if (!_booted.IsValueCreated) return;
            // KernelSupervisor.Dispose is itself idempotent and best-effort —
            // see its contract for the "no orphaned python.exe" guarantee.
            try { _booted.Value.Supervisor.Dispose(); } catch { /* best-effort */ }
        }
    }

    private Booted GetBooted()
    {
        if (Volatile.Read(ref _disposed) != 0)
            throw new ObjectDisposedException(nameof(KernelHost));
        lock (_gate)
        {
            var booted = _booted.Value;  // boots on first use, serialised by _gate
            if (KernelIsDead(booted.Supervisor))
            {
                // The kernel died — a crash, or a hard-cancel kill of a wedged
                // child. Dispose the husk and boot a fresh one so a caller never
                // hands work to a dead pipe (which would fail with a confusing
                // transport error on the next run).
                try { booted.Supervisor.Dispose(); } catch { /* best-effort */ }
                _booted = new Lazy<Booted>(Boot, LazyThreadSafetyMode.ExecutionAndPublication);
                booted = _booted.Value;
            }
            return booted;
        }
    }

    /// <summary>True when the supervisor's child process is gone (exited,
    /// killed, or disposed). Treated as dead on any probe failure so a husk is
    /// always replaced rather than reused.</summary>
    private static bool KernelIsDead(KernelSupervisor supervisor)
    {
        try { return supervisor.Process.HasExited; }
        catch { return true; }
    }

    private static Booted Boot()
    {
        // Resolve the active workbook directory once so both resolvers
        // give the per-project venv and the Setup-extracted
        // .pyexcel-kernel precedence over the PATH interpreter and the
        // bundled embedded/ copy. Null (unsaved/no workbook) falls back
        // to those bundled defaults. The kernel is process-wide and boots
        // lazily once, so this captures the workbook active at first use —
        // acceptable for the Phase-4 stop-gap host (see the type remarks).
        var key = PyExcelServices.WorkbookContext.CurrentWorkbookKey;
        var workbookDir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
        // Prefer the dedicated project directory the user chose on Enable (saved
        // in workbook state); otherwise fall back to the workbook-derived
        // default. A SharePoint/OneDrive-online workbook reports a URL, not a
        // local folder; ProjectDirectory maps that (and any other non-local
        // path) to the same local environment directory Setup provisions into,
        // so the runtime looks where Setup installed.
        var stored = key is null ? null : PyExcelServices.State.Get(key).ProjectDir;
        var projectDir = string.IsNullOrEmpty(stored)
            ? PyExcel.Common.ProjectDirectory.Resolve(workbookDir)
            : stored;
        var python = PythonResolver.ResolvePython(projectDir);
        var pythonPath = PythonResolver.ResolveEmbeddedPath(projectDir);
        var supervisor = KernelSupervisor.StartPython(python, pythonPath);
        var client = new KernelClient(supervisor);
#if NETFRAMEWORK
        ForwardKernelOutputToLog(supervisor, client);
#endif
        return new Booted(supervisor, client);
    }

#if NETFRAMEWORK
    /// <summary>
    /// Surface the kernel's runtime chatter in Excel-DNA's LogDisplay window
    /// so the user can actually see it: the subprocess's stdout/stderr (where
    /// user <c>print()</c> lands) and any structured LOG frames. Without this
    /// the output is captured but shown nowhere — the "Run Python didn't show
    /// the prints anywhere" report. Best-effort: a logging failure must never
    /// disturb a run, and these handlers live for the process lifetime
    /// alongside the singleton host.
    /// </summary>
    private static void ForwardKernelOutputToLog(KernelSupervisor supervisor, KernelClient client)
    {
        supervisor.OutputReceived += (_, e) =>
        {
            try
            {
                // Pass the line as the sole argument: LogDisplay.WriteLine only
                // runs string.Format when extra args follow, so brace-bearing
                // output (dict reprs, f-strings) can't throw FormatException.
                ExcelDna.Logging.LogDisplay.WriteLine(
                    (e.IsError ? "[python:stderr] " : "[python] ") + e.Text);
            }
            catch { /* never let logging disturb a run */ }
        };
        client.LogReceived += (_, e) =>
        {
            try { ExcelDna.Logging.LogDisplay.WriteLine($"[python:{e.Level}] {e.Text}"); }
            catch { /* never let logging disturb a run */ }
        };
    }
#endif

    private readonly struct Booted
    {
        public KernelSupervisor Supervisor { get; }
        public KernelClient Client { get; }

        public Booted(KernelSupervisor supervisor, KernelClient client)
        {
            Supervisor = supervisor;
            Client = client;
        }
    }
}
