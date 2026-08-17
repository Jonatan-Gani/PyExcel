"""Every kernel module must be embedded in the shipped .xll.

``PyExcel.Setup.csproj`` lists its ``<EmbeddedResource>`` items one file at a
time, and ``KernelResourceExtractor`` writes exactly those onto disk at Setup.
A module that exists in ``embedded/pyexcel`` but is missing from that list is
invisible in every test — ``tests/conftest.py`` puts ``embedded/`` on
``sys.path``, so the suite imports from the source tree and never from the
extracted copy — yet it is absent from the add-in a user actually installs.

That happened: ``declared_types.py`` shipped as source, was omitted from the
csproj, and the packed add-in could not run the typed contract at all while CI
stayed green. This test is the tripwire.
"""

import os
import re

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
CSPROJ = os.path.join(REPO_ROOT, "src", "PyExcel.Setup", "PyExcel.Setup.csproj")
KERNEL_PACKAGE = os.path.join(REPO_ROOT, "embedded", "pyexcel")


def _declared_logical_names():
    with open(CSPROJ, encoding="utf-8") as fh:
        return set(re.findall(r'LogicalName="(pyexcel/[^"]+)"', fh.read()))


def _python_modules_on_disk():
    modules = set()
    for root, _dirs, files in os.walk(KERNEL_PACKAGE):
        for name in files:
            if not name.endswith(".py"):
                continue
            relative = os.path.relpath(
                os.path.join(root, name), os.path.join(REPO_ROOT, "embedded")
            )
            modules.add(relative.replace(os.sep, "/"))
    return modules


def test_every_kernel_module_is_embedded():
    missing = _python_modules_on_disk() - _declared_logical_names()
    assert not missing, (
        "these kernel modules exist on disk but are not <EmbeddedResource> items "
        f"in PyExcel.Setup.csproj, so they will not ship inside the .xll: "
        f"{sorted(missing)}"
    )


def test_no_embedded_resource_points_at_a_missing_module():
    declared_py = {n for n in _declared_logical_names() if n.endswith(".py")}
    stale = declared_py - _python_modules_on_disk()
    assert not stale, (
        "PyExcel.Setup.csproj embeds these paths but no such module exists, "
        f"which fails the build at pack time: {sorted(stale)}"
    )


def test_the_typed_contract_module_is_embedded():
    """Named explicitly because the whole declared-type feature imports it."""
    assert "pyexcel/kernel/declared_types.py" in _declared_logical_names()
