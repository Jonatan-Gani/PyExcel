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
"""
