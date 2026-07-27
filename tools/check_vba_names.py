"""Static checks for VBA pitfalls that cause cryptic Excel compile/runtime errors.

1. Case-insensitive Dim/procedure name collisions (runtime error 91).
2. Statement-style Sub calls that pass a user-defined Function with arguments
   without ``Call`` (compile error: "Argument not optional").
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_FILE = ROOT / "vba" / "AssignedLateSchedule.bas"

DIM_RE = re.compile(r"^\s*Dim\s+(.+)$", re.IGNORECASE)
PROC_RE = re.compile(r"^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)", re.IGNORECASE)
FUNCTION_DEF_RE = re.compile(
    r"^\s*(?:Public|Private)?\s*Function\s+(\w+)", re.IGNORECASE
)


def _split_dim_names(dim_clause: str) -> list[str]:
    names: list[str] = []
    for part in dim_clause.split(","):
        token = part.strip().split(" ", 1)[0].strip()
        if token:
            names.append(token)
    return names


def find_collisions(text: str) -> list[str]:
    procedures: set[str] = set()
    variables: set[str] = set()

    for line in text.splitlines():
        fn = PROC_RE.match(line)
        if fn:
            procedures.add(fn.group(1).lower())
            continue
        dm = DIM_RE.match(line)
        if dm:
            for name in _split_dim_names(dm.group(1)):
                variables.add(name.lower())

    return sorted(variables & procedures)


def find_bad_sub_calls(text: str, user_functions: set[str]) -> list[str]:
    """Flag ``SubName arg, UserFunc(...)`` without Call.

    VBA mis-parses nested user-function calls in statement-style Sub invocations
    and raises compile error "Argument not optional".
    """
    joined = re.sub(r"_\s*\n\s*", " ", text)
    hits: list[str] = []
    # Match statement-style calls (no Call keyword) that include UserFunc(...).
    pattern = re.compile(
        r"(?im)(?:^|\bThen\s+)"
        r"(?!Call\b)"
        r"(?P<sub>[A-Za-z_]\w*)\b"
        r"(?!\s*=)"  # not an assignment
        r"[^\n]*?"
        r",\s*(?P<fn>[A-Za-z_]\w*)\s*\("
    )
    for match in pattern.finditer(joined):
        sub = match.group("sub")
        fn = match.group("fn")
        if sub.lower() in {"if", "for", "do", "while", "select", "with", "elseif"}:
            continue
        if fn.lower() not in user_functions:
            continue
        # Ignore if this was already a Call statement (Call SubName(...)).
        window_start = max(0, match.start() - 10)
        prefix = joined[window_start : match.start()]
        if re.search(r"\bCall\s+$", prefix, re.I):
            continue
        hits.append(
            f"{sub} ..., {fn}(...)  -> use Call {sub}(..., {fn}(...))"
        )
    return hits


def main() -> int:
    text = VBA_FILE.read_text(encoding="utf-8")
    user_functions = {
        m.group(1).lower() for m in FUNCTION_DEF_RE.finditer(text)
    }

    collisions = find_collisions(text)
    bad_calls = find_bad_sub_calls(text, user_functions)
    failed = False

    if collisions:
        failed = True
        print("VBA name collisions detected (case-insensitive):")
        for name in collisions:
            print(f"  - {name!r}: Dim shadows Sub/Function")

    if bad_calls:
        failed = True
        print('Risky Sub calls without Call (can cause "Argument not optional"):')
        for hit in bad_calls:
            print(f"  - {hit}")

    if failed:
        return 1

    print(f"VBA static checks passed for {VBA_FILE.name}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
