"""Compile the VBA source into ``assets/vbaProject.bin`` (no Excel required).

XlsxWriter can embed a ``vbaProject.bin`` but cannot compile VBA.  Historically
that binary had to be exported from a copy of Excel.  This script removes that
manual step: it uses the pure-Python ``pyOpenVBA`` library to inject the module
in ``vba/AssignedLateSchedule.bas`` into a macro-enabled workbook and writes the
resulting project back to ``assets/vbaProject.bin``.

Run this whenever ``vba/AssignedLateSchedule.bas`` changes:

    python tools/compile_vba.py

Then run ``python build_workbook.py`` to produce the shipped workbook with the
refreshed macros.  Excel recompiles the p-code from this source on first open.
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

VBA_SOURCE = ROOT / "vba" / "AssignedLateSchedule.bas"
VBA_PROJECT = ROOT / "assets" / "vbaProject.bin"
WORKBOOK = ROOT / "Assigned Late Schedule.xlsm"
MODULE_NAME = "AssignedLateSchedule"
# XlsxWriter's example project ships these demo/duplicate modules; drop them.
DEMO_MODULES = ("Module1", "ThisWorkbook1")


def _load_source() -> str:
    text = VBA_SOURCE.read_text(encoding="utf-8")
    # VBA module streams use CRLF line endings.
    return text.replace("\r\n", "\n").replace("\n", "\r\n")


def compile_vba() -> None:
    from pyopenvba import ExcelFile, VBAModuleKind

    import build_workbook

    # A macro-enabled seed workbook is required to host the project.  Building
    # embeds the current assets/vbaProject.bin; injection below is idempotent so
    # it works whether that binary is the real project or a placeholder.
    build_workbook.build_workbook()

    code = _load_source()
    wb = ExcelFile(str(WORKBOOK))
    project = wb.vba_project()

    if MODULE_NAME in wb.module_names():
        wb.set_module(MODULE_NAME, code)
    else:
        project.add_module(MODULE_NAME, code, kind=VBAModuleKind.standard)

    for stray in DEMO_MODULES:
        if stray in wb.module_names():
            project.delete_module(stray)

    problems = wb.validate()
    if problems:
        raise SystemExit(f"VBA project failed validation: {problems}")

    wb.save(str(WORKBOOK))
    wb.close()

    with zipfile.ZipFile(str(WORKBOOK)) as archive:
        VBA_PROJECT.write_bytes(archive.read("xl/vbaProject.bin"))

    print(f"Wrote {VBA_PROJECT.relative_to(ROOT)} with module '{MODULE_NAME}'.")
    print("Now run: python build_workbook.py")


if __name__ == "__main__":
    compile_vba()
