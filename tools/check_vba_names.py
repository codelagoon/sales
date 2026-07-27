"""Fail if a VBA module declares a variable with the same name as a procedure.

VBA is case-insensitive, so ``Dim recentCounts`` shadows ``RecentCounts`` and
calling the function from the same module raises runtime error 91.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_FILE = ROOT / "vba" / "AssignedLateSchedule.bas"

DIM_RE = re.compile(r"^\s*Dim\s+(.+)$", re.IGNORECASE)
FUNCTION_RE = re.compile(r"^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)", re.IGNORECASE)
SUB_RE = FUNCTION_RE


def _split_dim_names(dim_clause: str) -> list[str]:
    names: list[str] = []
    for part in dim_clause.split(","):
        token = part.strip().split(" ", 1)[0].strip()
        if token:
            names.append(token)
    return names


def find_collisions(text: str) -> list[tuple[str, str, str]]:
    procedures: set[str] = set()
    variables: set[str] = set()

    for line in text.splitlines():
        fn = FUNCTION_RE.match(line)
        if fn:
            procedures.add(fn.group(1).lower())
            continue
        dm = DIM_RE.match(line)
        if dm:
            for name in _split_dim_names(dm.group(1)):
                variables.add(name.lower())

    collisions = []
    for var in sorted(variables):
        if var in procedures:
            collisions.append((var, "variable", "procedure"))
    return collisions


def main() -> int:
    text = VBA_FILE.read_text(encoding="utf-8")
    collisions = find_collisions(text)
    if collisions:
        print("VBA name collisions detected (case-insensitive):")
        for name, kind_a, kind_b in collisions:
            print(f"  - {name!r}: {kind_a} shadows {kind_b}")
        return 1
    print(f"No VBA name collisions in {VBA_FILE.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
