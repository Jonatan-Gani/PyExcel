"""PyExcel v2 kernel: persistent supervisor + worker model.

Phase 1 ships only the framing primitive. Subsequent phases add:

* ``supervisor.py`` — accepts pipe connections from the C# client, dispatches
  Run requests to a worker process, watchdogs the parent PID.
* ``worker.py`` — imports user scripts, invokes their ``@job``-decorated
  functions, returns Arrow IPC payloads.
* ``transport.py`` — Windows named-pipe server-side wrapper.

The framing protocol is the lowest layer everything else builds on; it is
intentionally pure stdlib and bounded so a malformed peer cannot crash us.
"""
