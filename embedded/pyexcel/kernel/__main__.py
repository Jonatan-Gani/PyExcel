"""Entry point for ``python -m pyexcel.kernel``.

The C# :class:`KernelSupervisor` spawns this module with ``--pipe <name>``
and then waits on its server-side pipe for the connection-back. Keep this
file tiny — the actual work lives in :mod:`pyexcel.kernel.supervisor`.
"""

from pyexcel.kernel.supervisor import main

if __name__ == "__main__":
    main()
