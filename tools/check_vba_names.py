"""Static checks for VBA pitfalls that cause cryptic Excel compile/runtime errors.

1. Case-insensitive Dim/procedure name collisions (runtime error 91).
2. Sub calls without ``Call`` that pass a user function with arguments
   (compile error: "Argument not optional").
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_FILE = ROOT / "vba" / "AssignedLateSchedule.bas"

DIM_RE = re.compile(r"^\s*Dim\s+(.+)$", re.IGNORECASE)
FUNCTION_RE = re.compile(r"^\s*(?:Public|Private)?\s*(?:Sub|Function)\s+(\w+)", re.IGNORECASE)

# Statement-style Sub call whose argument list contains FuncName(...).
# Example that fails to compile:  Increment dict, PairKey(a, b)
BAD_SUB_CALL_RE = re.compile(
    r"(?im)^(?!\s*(?:If\b.*\bThen\s+)?Call\b)"  # not already using Call
    r"(?:\s*(?:If\b.+\bThen))?.*"
    r"\b(?P<sub>Increment|LoadHistoryStatistics|AssignSlotsGlobally|ImproveSchedule|"
    r"WriteSlotsToSchedule|CommitAssignmentToHistory|ApplyTemporaryAssignment|"
    r"GetCountSpreadAfter|RemoveUnusedFutureBlankDates|SortScheduleByDate|"
    r"SetNamedValue|CheckActiveStaffThreshold)\b"
    r"[^(=\n]*,\s*"
    r"(?P<fn>[A-Za-z_]\w*)\s*\("
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
        fn = FUNCTION_RE.match(line)
        if fn:
            procedures.add(fn.group(1).lower())
            continue
        dm = DIM_RE.match(line)
        if dm:
            for name in _split_dim_names(dm.group(1)):
                variables.add(name.lower())

    return sorted(variables & procedures)


def find_bad_sub_calls(text: str) -> list[str]:
    # Join continued lines so "Then _\n    Increment ..." is visible as one statement.
    joined = re.sub(r"_\s*\n\s*", " ", text)
    hits = []
    for match in BAD_SUB_CALL_RE.finditer(joined):
        # Skip if the function name is actually Call already handled; keep report short.
        snippet = match.group(0).strip()
        if re.search(r"\bCall\s+" + re.escape(match.group("sub")), snippet, re.I):
            continue
        hits.append(
            f"{match.group('sub')} ..., {match.group('fn')}(...)  -> use Call {match.group('sub')}(..., {match.group('fn')}(...))"
        )
    return hits


def main() -> int:
    text = VBA_FILE.read_text(encoding="utf-8")
    collisions = find_collisions(text)
    bad_calls = find_bad_sub_calls(text)
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
