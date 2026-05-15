"""
Pytest configuration for PyExcel v2.

Adds ``src/embedded`` to ``sys.path`` so the ``pyexcel`` package is importable
without installing it. This matches how the kernel will be invoked at runtime
(``python -m pyexcel.kernel`` with ``PYTHONPATH`` set by ``python.bas``).
"""

from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
_EMBEDDED = _REPO_ROOT / "src" / "embedded"

if str(_EMBEDDED) not in sys.path:
    sys.path.insert(0, str(_EMBEDDED))
