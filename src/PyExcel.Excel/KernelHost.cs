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

    private readonly Lazy<Booted> _booted;
    private int _disposed;

    public KernelHost()
    {
        _booted = new Lazy<Booted>(Boot, LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>The typed client. First access boots the kernel.</summary>
    public KernelClient Client => GetBooted().Client;

    /// <summary>The supervisor (process + pipe owner).</summary>
    public KernelSupervisor Supervisor => GetBooted().Supervisor;

    /// <summary>True once the kernel has been booted at least once.</summary>
    public bool IsStarted => _booted.IsValueCreated;

    /// <summary>
    /// Idempotent shutdown. Safe to call multiple times; safe to call
    /// before the kernel was ever booted (no-op in that case).
    /// </summary>
    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        if (!_booted.IsValueCreated) return;
        // KernelSupervisor.Dispose is itself idempotent and best-effort —
        // see its contract for the "no orphaned python.exe" guarantee.
        _booted.Value.Supervisor.Dispose();
    }

    private Booted GetBooted()
    {
        if (Volatile.Read(ref _disposed) != 0)
            throw new ObjectDisposedException(nameof(KernelHost));
        return _booted.Value;
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
        var workbookDir = PyExcelServices.WorkbookContext.CurrentWorkbookDirectory;
        var python = PythonResolver.ResolvePython(workbookDir);
        var pythonPath = PythonResolver.ResolveEmbeddedPath(workbookDir);
        var supervisor = KernelSupervisor.StartPython(python, pythonPath);
        var client = new KernelClient(supervisor);
        return new Booted(supervisor, client);
    }

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
