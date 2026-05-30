using System;
using System.Collections.Generic;

namespace PyExcel.State;

/// <summary>
/// Per-workbook "last error" registry. Tracks one
/// <see cref="KernelErrorRecord"/> per workbook key plus a global slot
/// for errors that arrive before any workbook is bound (the
/// <c>=PY.RUN</c> UDF when called from a fresh sheet, the kernel boot
/// failure, …). The ribbon reads this to decide whether to enable the
/// Show / Copy Last Error buttons, and to render the message for them.
///
/// <para>Mirrors <see cref="StateService"/>'s threading model: single
/// coarse lock around the dictionary, a synchronous
/// <see cref="ErrorChanged"/> event fired after each mutation so the
/// ribbon can invalidate. The lock is never held while the event runs.</para>
/// </summary>
public sealed class ErrorService
{
    private readonly Dictionary<string, KernelErrorRecord> _byWorkbook =
        new(StringComparer.Ordinal);
    private KernelErrorRecord? _global;
    private readonly object _lock = new();

    /// <summary>Fired after each <see cref="Record"/> or
    /// <see cref="Clear"/>. The argument is the affected workbook key
    /// (or <see langword="null"/> for the global slot).</summary>
    public event EventHandler<ErrorChangedEventArgs>? ErrorChanged;

    /// <summary>
    /// Stash one error. A null <paramref name="workbookKey"/> targets
    /// the global slot — used when no workbook is active at the time of
    /// the failure (kernel boot, add-in init).
    /// </summary>
    public void Record(string? workbookKey, KernelErrorRecord record)
    {
        if (record is null) throw new ArgumentNullException(nameof(record));

        lock (_lock)
        {
            if (workbookKey is null) _global = record;
            else _byWorkbook[workbookKey] = record;
        }
        ErrorChanged?.Invoke(this, new ErrorChangedEventArgs(workbookKey));
    }

    /// <summary>
    /// Return the last error for a workbook, falling back to the global
    /// slot if the workbook has none. Returns <see langword="null"/> if
    /// neither slot has anything — that's how the ribbon decides to
    /// disable the Show / Copy buttons.
    /// </summary>
    public KernelErrorRecord? GetLast(string? workbookKey)
    {
        lock (_lock)
        {
            if (workbookKey is not null
                && _byWorkbook.TryGetValue(workbookKey, out var r))
                return r;
            return _global;
        }
    }

    /// <summary>
    /// Drop the stored error for a workbook (or the global slot, when
    /// <paramref name="workbookKey"/> is <see langword="null"/>). No-op
    /// if the slot was already empty — but always fires
    /// <see cref="ErrorChanged"/> so the ribbon can repaint without a
    /// separate "had-it" check.
    /// </summary>
    public void Clear(string? workbookKey)
    {
        lock (_lock)
        {
            if (workbookKey is null) _global = null;
            else _byWorkbook.Remove(workbookKey);
        }
        ErrorChanged?.Invoke(this, new ErrorChangedEventArgs(workbookKey));
    }
}

/// <summary>Argument for <see cref="ErrorService.ErrorChanged"/>.</summary>
public sealed class ErrorChangedEventArgs : EventArgs
{
    /// <summary>The affected workbook key, or <see langword="null"/>
    /// for the global slot.</summary>
    public string? WorkbookKey { get; }

    public ErrorChangedEventArgs(string? workbookKey)
    {
        WorkbookKey = workbookKey;
    }
}
