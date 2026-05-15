"""Pytest configuration for PyExcel v2.

Adds ``embedded/`` to ``sys.path`` so the ``pyexcel`` package is importable
without installing it. Matches how the kernel is invoked at runtime:
``python -m pyexcel.kernel`` against an extracted ``embedded/`` directory
inside the project's venv-adjacent folder.
"""

from __future__ import annotations

import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
_EMBEDDED = _REPO_ROOT / "embedded"

if str(_EMBEDDED) not in sys.path:
    sys.path.insert(0, str(_EMBEDDED))
