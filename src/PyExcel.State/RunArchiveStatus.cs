namespace PyExcel.State;

/// <summary>
/// Outcome of a single archived run. <see cref="Cancelled"/> is split out
/// from <see cref="Error"/> because the kernel reports cancellation
/// distinctly — see <c>KernelException</c> with code <c>"Cancelled"</c> and
/// the supervisor's CANCEL frame handling. Users replaying an archive want
/// to know whether the absence of an output means "the script errored" or
/// "the user pulled the rug while it was running".
/// </summary>
public enum RunArchiveStatus
{
    /// <summary>The kernel returned a result (possibly <c>None</c>).</summary>
    Success,
    /// <summary>The run failed — either a kernel-side <c>KernelException</c>
    /// or a host-side fault during marshalling / range I/O.</summary>
    Error,
    /// <summary>The run was cancelled before the kernel produced a result —
    /// formula change, workbook close, or an explicit Cancel from the UI.</summary>
    Cancelled,
}
