# kernel (tests)

## Macro
The pytest suite for the Python kernel (`embedded/pyexcel/kernel`). Unit tests
exercise the pure modules directly; `test_supervisor` is an end-to-end test that
spawns a real `python -m pyexcel.kernel` subprocess and plays the C# server role
over an `AF_UNIX` socket. `conftest.py` (in `tests/`) puts `embedded/` on the
path. Run with `pytest tests/`.

## Files
### test_framing.py
Covers `framing.py`: frame encode/decode round-trips, size-bound enforcement, and
malformed/truncated-frame handling. Inputs: constructed frames and byte buffers. Output:
assertions (no artifacts).

### test_arrow_io.py
Covers `arrow_io.py`: shape-preserving encode/decode of DataFrame/list/scalar values and
the `types` wrappers. Inputs: Python values and Arrow buffers. Output: assertions.

### test_chart.py
Covers `chart.py`: Plotly-figure→`ChartSpec` and Matplotlib-figure→`ChartImage`
conversion, including duck-typed detection and unsupported-type errors. Inputs: figure
stand-ins. Output: assertions.

### test_cross_language_vectors.py
Verifies the Python encoding matches the C# side byte-for-byte against shared fixture
vectors (the wire-compatibility contract). Inputs: shared test vectors. Output:
assertions.

### test_worker.py
Covers `worker.py`: `run_job` shape coverage and error codes, cooperative
cancellation/progress, and the `input()`/stdin guard. Inputs: written user scripts and
Arrow payloads. Output: assertions.

### test_supervisor.py
End-to-end test of the supervisor loop: spawns a real kernel subprocess and drives
HELLO/PING/RUN_REQUEST/CANCEL/PROGRESS/SHUTDOWN over a Unix-domain socket, including
prompt cancellation of a non-cooperative worker. Inputs: a spawned kernel + crafted
frames. Output: assertions (skipped on Windows).

### __init__.py
Empty package marker. Inputs/Output: none.
