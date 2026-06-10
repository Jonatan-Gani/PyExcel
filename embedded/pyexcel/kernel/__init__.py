"""PyExcel v2 kernel: persistent supervisor + worker model.

Modules:

* :mod:`pyexcel.kernel.framing` — the wire format (length-prefixed binary
  frames + canonical JSON meta). Pure stdlib, bounded against malformed
  peers, mirrored byte-for-byte by ``PyExcel.Bridge/Framing.cs``.
* :mod:`pyexcel.kernel.transport` — POSIX (AF_UNIX) / Windows (named pipe)
  client wrappers. The C# ``KernelSupervisor`` owns the pipe server; we
  connect into it.
* :mod:`pyexcel.kernel.arrow_io` — Arrow IPC marshalling for the data
  plane. Encodes DataFrame / list / scalar payloads with shape metadata
  so the host can spill back into the right cell geometry.
* :mod:`pyexcel.kernel.chart` — figure conversion. A ``transform()``
  returning a Plotly figure ships a JSON chart spec the host renders as
  a native Excel chart; a Matplotlib figure ships rendered image bytes
  (SVG, PNG fallback) the host embeds as a picture.
* :mod:`pyexcel.kernel.worker` — pure ``run_job(meta, payloads) -> JobOutcome``.
  Loads the user's script, calls the target function, marshals the result
  back. Caches user modules by absolute path + mtime so unchanged scripts
  don't re-import on every call.
* :mod:`pyexcel.kernel.supervisor` — the in-process event loop. Performs
  the HELLO handshake, then dispatches PING/PONG, RUN_REQUEST (via
  :mod:`worker`), and SHUTDOWN.
* :mod:`pyexcel.kernel.__main__` — the ``python -m pyexcel.kernel`` entry
  point. Parses argv, opens the transport, hands control to the
  supervisor.

User-facing helpers
-------------------

Long-running transform functions can poll :func:`is_cancelled` between
work units and return early; the supervisor will then surface an
``ERROR`` frame with code ``"Cancelled"`` instead of ``RUN_RESULT``.
They can also call :func:`report_progress` to push status updates to the
host as ``PROGRESS`` frames (driving the progress UI). Both are safe to
call unconditionally — they're inert no-ops when no job is in flight.
"""

from .types import ChartImage, ChartSpec, Formula  # noqa: E402, F401 — re-export for user scripts
from .worker import is_cancelled, report_progress  # noqa: E402, F401 — re-export for user scripts

__all__ = ["ChartImage", "ChartSpec", "Formula", "is_cancelled", "report_progress"]
