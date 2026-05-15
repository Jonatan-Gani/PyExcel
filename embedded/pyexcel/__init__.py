"""PyExcel v2 Python kernel.

This package is shipped *inside* the .xll as an embedded resource and
extracted to the project's .venv-adjacent folder on first Setup. The C#
side spawns it as a subprocess via ``python -m pyexcel.kernel``.

See /docs/v2-build.md for the runtime project layout.
"""

__version__ = "2.0.0a0"
