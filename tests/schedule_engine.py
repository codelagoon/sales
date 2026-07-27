"""Faithful Python port of the VBA assignment/fairness engine.

This module mirrors the logic in ``vba/AssignedLateSchedule.bas`` so the
scheduling and fairness behaviour can be exercised and verified with automated
tests *before* the VBA is compiled into the workbook.  Every scoring term,
weight, tie-break and constraint is kept identical to the VBA so the tests are
a meaningful proxy for the real macro.

Keep this file in sync with the .bas module.  If you change the algorithm in
one place, change it in both and re-run ``tests/test_schedule.py``.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Dict, List, Optional, Tuple

LOCATION_ONE = "1 Centre Street"
LOCATION_TWO = "100 Centre Street"
PAIR_SEPARATOR = " / "

WEEKDAY_NUMBERS = {
    "monday": 1,
    "tuesday": 2,
    "wednesday": 3,
    "thursday": 4,
    "friday": 5,
    "saturday": 6,
    "sunday": 7,
}


@dataclass
class Slot:
    schedule_date: date
    location: str
    table_row: int
    table_column: int
    mechanic: str = ""
    note: str = ""


@dataclass
class HistoryRow:
    date: date
    location: str
    mechanic: str
    created_on: object = None
    note: str = ""


@dataclass
class Workbook:
    """In-memory stand-in for the four worksheets/tables."""

    mechanics: List[Tuple[str, str]] = field(default_factory=list)  # (name, active)
    availability: List[Tuple[str, date, str]] = field(default_factory=list)
    history: List[HistoryRow] = field(default_factory=list)
    # schedule rows: dict with keys 'date', LOCATION_ONE, LOCATION_TWO
    schedule: List[dict] = field(default_factory=list)


# --- weekday helpers (mirror WeekdayNumber / FirstMatchingDate) --------------


def weekday_number(name: str) -> int:
    key = name.strip().lower()
    if key not in WEEKDAY_NUMBERS:
        raise ValueError("Choose a valid Day of Week.")
    return WEEKDAY_NUMBERS[key]


def _weekday_monday(d: date) -> int:
    # VBA Weekday(d, vbMonday): Monday=1 .. Sunday=7. Python weekday(): Mon=0.
    return d.weekday() + 1


def first_matching_date(start_date: date, requested_day: int) -> date:
    offset = (requested_day - _weekday_monday(start_date) + 7) % 7
    return start_date + timedelta(days=offset)


# --- date generation (mirror GenerateDatesInternal) --------------------------


def generate_dates(wb: Workbook, start_date: date, number_of_weeks: int, day_of_week: str,
                   today: date) -> None:
    if number_of_weeks < 1:
        raise ValueError("Number of Weeks must be at least 1.")
    requested_day = weekday_number(day_of_week)

    # Build the exact future date set implied by the settings, then replace any
    # future Schedule rows that are not in that set (wrong weekday, outside the
    # week window, leftover blanks). Past dates and History are untouched.
    target_dates = {}
    generated = first_matching_date(start_date, requested_day)
    for _ in range(number_of_weeks):
        if generated >= today:
            target_dates[_date_key(generated)] = generated
        generated = generated + timedelta(days=7)

    for row in wb.schedule:
        d = row.get("date")
        if _is_future(d, today) and _date_key(d) not in target_dates:
            row["date"] = None
            row[LOCATION_ONE] = ""
            row[LOCATION_TWO] = ""
        elif _blank(d):
            row[LOCATION_ONE] = ""
            row[LOCATION_TWO] = ""

    existing = {_date_key(r["date"]) for r in wb.schedule if isinstance(r.get("date"), date)}
    for generated in target_dates.values():
        if _date_key(generated) not in existing:
            empty = _first_empty_row(wb)
            if empty is None:
                empty = {"date": None, LOCATION_ONE: "", LOCATION_TWO: ""}
                wb.schedule.append(empty)
            empty["date"] = generated
            existing.add(_date_key(generated))

    wb.schedule = [r for r in wb.schedule if isinstance(r.get("date"), date)]
    wb.schedule.sort(key=lambda r: r["date"])


# --- assignment engine (mirror AssignFutureSchedule) -------------------------


def assign_future_schedule(wb: Workbook, today: date, recent_cutoff: Optional[date] = None) -> List[Slot]:
    eligible = active_mechanics(wb)
    if len(eligible) < 4:
        raise ValueError(
            "At least four active mechanics are required "
            "(two locations x two mechanics per schedule date).")

    # Operational safeguard: refuse to run if any generated date cannot be
    # staffed with enough active, available mechanics.
    check_active_staff_threshold(wb, today)

    slots = _collect_empty_future_slots(wb, today)
    if not slots:
        return []

    if recent_cutoff is None:
        recent_cutoff = _add_months(today, -3)

    total_counts = _lifetime_counts(wb, eligible)
    recent_counts = _recent_counts(wb, eligible, recent_cutoff)
    location_counts = _lifetime_location_counts(wb, eligible)
    last_assigned = _last_assignment_dates(wb, eligible)
    pair_counts = _lifetime_pair_counts(wb)

    _assign_slots_globally(slots, eligible, total_counts, recent_counts,
                           location_counts, last_assigned, pair_counts, wb)
    _improve_schedule(slots, eligible, wb)
    _write_slots_to_schedule(wb, slots)
    _commit_assignment_to_history(wb, slots)
    return slots


def check_active_staff_threshold(wb: Workbook, today: date) -> None:
    """Abort with a clear message if any future date lacks enough available staff.

    Every empty location cell needs two mechanics, so a fully empty date needs
    four distinct active mechanics who are not marked unavailable that day.
    """
    eligible = active_mechanics(wb)
    for row in wb.schedule:
        d = row.get("date")
        if not _is_future(d, today):
            continue
        needed = 0
        for colname in (LOCATION_ONE, LOCATION_TWO):
            if _blank(row.get(colname)):
                needed += 2
        if needed == 0:
            continue
        available = [m for m in eligible if not _is_unavailable(wb, m, d)]
        if len(available) < needed:
            raise ValueError(
                f"Only {len(available)} active, available mechanic(s) on "
                f"{d:%m/%d/%Y}; {needed} are required. Add active mechanics or "
                f"remove availability blocks for that date, then try again.")


def _assign_slots_globally(slots, eligible, total_counts, recent_counts, location_counts,
                           last_assigned, pair_counts, wb) -> None:
    for _ in range(len(slots)):
        best_score = 1e30
        best_slot = None
        best_mechanic = ""
        target_date = _next_unassigned_date(slots)
        for cand_slot in slots:
            if cand_slot.mechanic == "" and cand_slot.schedule_date == target_date:
                for candidate in eligible:
                    if _is_eligible_for_slot(candidate, cand_slot, slots, wb):
                        score = _assignment_score(candidate, cand_slot, slots, total_counts,
                                                  recent_counts, location_counts, last_assigned,
                                                  pair_counts)
                        if score < best_score or (score == best_score and candidate < best_mechanic):
                            best_score = score
                            best_slot = cand_slot
                            best_mechanic = candidate
        if best_slot is None:
            raise ValueError("No eligible mechanic is available for every future schedule date.")
        # Explicit duplicate-assignment guard: never place a mechanic at both
        # buildings on the same date (defensive; scoring already avoids it).
        if _creates_same_day_duplicate(best_mechanic, slots.index(best_slot), slots):
            raise ValueError(
                f"{best_mechanic} would be assigned twice on {best_slot.schedule_date:%m/%d/%Y}.")
        best_slot.note = _build_selection_note(best_mechanic, best_slot, total_counts,
                                               recent_counts, location_counts, last_assigned,
                                               pair_counts, slots)
        best_slot.mechanic = best_mechanic
        _apply_temporary_assignment(best_mechanic, best_slot, total_counts, recent_counts,
                                    location_counts, last_assigned, pair_counts, slots)


def _assignment_score(mechanic, slot, slots, total_counts, recent_counts, location_counts,
                      last_assigned, pair_counts) -> float:
    lifetime = total_counts[mechanic]
    recent = recent_counts[mechanic]
    days_since = _days_since_last_assignment(last_assigned[mechanic], slot.schedule_date)
    location_gap = _location_imbalance_after(mechanic, slot.location, location_counts)
    maximum_after, minimum_after = _count_spread_after(mechanic, total_counts)
    consecutive_penalty = _consecutive_date_penalty(mechanic, slot.schedule_date, slots)
    pair_penalty = _pair_repeat_penalty(mechanic, slot, slots, pair_counts)

    return ((maximum_after - minimum_after) * 1_000_000_000.0
            + lifetime * 1_000_000.0
            + recent * 10_000.0
            + consecutive_penalty * 1_000.0
            + pair_penalty * 100.0
            + location_gap * 10.0
            - min(days_since, 9999) / 100_000.0)


def _is_eligible_for_slot(mechanic, slot, slots, wb) -> bool:
    if _is_unavailable(wb, mechanic, slot.schedule_date):
        return False
    for s in slots:
        if s.schedule_date == slot.schedule_date and s.mechanic == mechanic:
            return False
    return True


def _apply_temporary_assignment(mechanic, slot, total_counts, recent_counts, location_counts,
                                last_assigned, pair_counts, slots) -> None:
    total_counts[mechanic] += 1
    recent_counts[mechanic] += 1
    location_counts[mechanic + "|" + slot.location] = location_counts.get(mechanic + "|" + slot.location, 0) + 1
    last_assigned[mechanic] = slot.schedule_date
    partner = _assigned_partner(slot, slots)
    if partner:
        key = _pair_key(mechanic, partner)
        pair_counts[key] = pair_counts.get(key, 0) + 1


def _improve_schedule(slots, eligible, wb) -> None:
    # Seed the improvement objective with lifetime location history so swaps
    # balance each mechanic across both locations over time, not just inside
    # the current batch.
    base_locations = _lifetime_location_counts(wb, eligible)
    changed = True
    while changed:
        changed = False
        for i in range(len(slots)):
            for j in range(i + 1, len(slots)):
                first_m = slots[i].mechanic
                second_m = slots[j].mechanic
                if first_m == second_m:
                    continue
                if (not _is_unavailable(wb, first_m, slots[j].schedule_date)
                        and not _is_unavailable(wb, second_m, slots[i].schedule_date)
                        and not _creates_same_day_duplicate(first_m, j, slots)
                        and not _creates_same_day_duplicate(second_m, i, slots)):
                    before = _local_fairness_score(slots, base_locations)
                    slots[i].mechanic, slots[j].mechanic = second_m, first_m
                    after = _local_fairness_score(slots, base_locations)
                    if after < before:
                        changed = True
                    else:
                        slots[i].mechanic, slots[j].mechanic = first_m, second_m


def _local_fairness_score(slots, base_locations) -> float:
    totals: Dict[str, int] = {}
    pairs: Dict[str, int] = {}
    locations: Dict[str, int] = dict(base_locations)
    score = 0.0
    for s in slots:
        totals[s.mechanic] = totals.get(s.mechanic, 0) + 1
        locations[s.mechanic + "|" + s.location] = locations.get(s.mechanic + "|" + s.location, 0) + 1
        partner = _assigned_partner(s, slots)
        if partner:
            key = _pair_key(s.mechanic, partner)
            pairs[key] = pairs.get(key, 0) + 1
        score += _consecutive_date_penalty(s.mechanic, s.schedule_date, slots) * 1000.0
    max_count = max(totals.values())
    min_count = min(totals.values())
    for key in pairs:
        score += pairs[key] * pairs[key] * 100.0
    # Location imbalance: weighted below pairing (100) per the spec priority.
    mechanics_seen = {k.rsplit("|", 1)[0] for k in locations}
    for mech in mechanics_seen:
        one = locations.get(mech + "|" + LOCATION_ONE, 0)
        two = locations.get(mech + "|" + LOCATION_TWO, 0)
        score += abs(one - two) * 10.0
    return score + (max_count - min_count) * 1_000_000_000.0


# --- write-out + history (mirror WriteSlotsToSchedule/AppendScheduleToHistory)


def _write_slots_to_schedule(wb, slots) -> None:
    grouped: Dict[str, str] = {}
    for s in slots:
        key = f"{s.table_row}|{s.table_column}"
        if key in grouped:
            grouped[key] = grouped[key] + PAIR_SEPARATOR + s.mechanic
        else:
            grouped[key] = s.mechanic
    for key, value in grouped.items():
        row_idx, col = key.split("|")
        row = wb.schedule[int(row_idx)]
        colname = LOCATION_ONE if col == "1" else LOCATION_TWO
        if _blank(row.get(colname)):
            row[colname] = value


def _commit_assignment_to_history(wb, slots) -> None:
    """Append new assignments to permanent history with an audit note that
    records why each mechanic was selected."""
    existing = set()
    for h in wb.history:
        existing.add(f"{_date_key(h.date)}|{h.location}|{h.mechanic}")
    for s in slots:
        key = f"{_date_key(s.schedule_date)}|{s.location}|{s.mechanic}"
        if key not in existing:
            wb.history.append(HistoryRow(s.schedule_date, s.location, s.mechanic, "now", s.note))
            existing.add(key)


def _build_selection_note(mechanic, slot, total_counts, recent_counts, location_counts,
                          last_assigned, pair_counts, slots) -> str:
    """Short audit tag explaining why this mechanic won the slot."""
    lifetime = total_counts[mechanic]
    recent = recent_counts[mechanic]
    last = last_assigned[mechanic]
    idle = "new" if last <= date(1900, 1, 1) else f"{max(0, (slot.schedule_date - last).days)}d"
    partner = _assigned_partner(slot, slots)
    if partner:
        repeats = pair_counts.get(_pair_key(mechanic, partner), 0)
        pair_txt = f"new pairing w/ {partner}" if repeats == 0 else f"paired w/ {partner} x{repeats}"
    else:
        pair_txt = "pairing pending"
    loc_gap = _location_imbalance_after(mechanic, slot.location, location_counts)
    return (f"lifetime {lifetime}, recent {recent}, idle {idle}, "
            f"{pair_txt}, loc-gap {loc_gap}")


# --- statistics (mirror the History* helpers) --------------------------------


def active_mechanics(wb) -> List[str]:
    result = []
    for name, active in wb.mechanics:
        if str(active).strip().lower() == "yes" and name.strip():
            result.append(name.strip())
    return result


def _lifetime_counts(wb, mechanics) -> Dict[str, int]:
    result = {m: 0 for m in mechanics}
    for h in wb.history:
        if h.mechanic in result:
            result[h.mechanic] += 1
    return result


def _recent_counts(wb, mechanics, cutoff) -> Dict[str, int]:
    result = {m: 0 for m in mechanics}
    for h in wb.history:
        if h.mechanic in result and h.date >= cutoff:
            result[h.mechanic] += 1
    return result


def _lifetime_location_counts(wb, mechanics) -> Dict[str, int]:
    result = {}
    for m in mechanics:
        result[m + "|" + LOCATION_ONE] = 0
        result[m + "|" + LOCATION_TWO] = 0
    for h in wb.history:
        key = h.mechanic + "|" + h.location
        if key in result:
            result[key] += 1
    return result


def _last_assignment_dates(wb, mechanics) -> Dict[str, date]:
    result = {m: date(1900, 1, 1) for m in mechanics}
    for h in wb.history:
        if h.mechanic in result and isinstance(h.date, date):
            if h.date > result[h.mechanic]:
                result[h.mechanic] = h.date
    return result


def _lifetime_pair_counts(wb) -> Dict[str, int]:
    grouped: Dict[str, List[str]] = {}
    for h in wb.history:
        if isinstance(h.date, date) and h.mechanic.strip():
            key = f"{_date_key(h.date)}|{h.location}"
            grouped.setdefault(key, []).append(h.mechanic.strip())
    result: Dict[str, int] = {}
    for people in grouped.values():
        if len(people) == 2:
            k = _pair_key(people[0], people[1])
            result[k] = result.get(k, 0) + 1
    return result


# --- low-level helpers -------------------------------------------------------


def _collect_empty_future_slots(wb, today) -> List[Slot]:
    slots: List[Slot] = []
    for row_index, row in enumerate(wb.schedule):
        d = row.get("date")
        if _is_future(d, today):
            for col, colname in ((1, LOCATION_ONE), (2, LOCATION_TWO)):
                if _blank(row.get(colname)):
                    for _ in range(2):
                        slots.append(Slot(d, colname, row_index, col))
    return slots


def _next_unassigned_date(slots) -> Optional[date]:
    earliest = None
    for s in slots:
        if s.mechanic == "":
            if earliest is None or s.schedule_date < earliest:
                earliest = s.schedule_date
    return earliest


def _is_unavailable(wb, mechanic, schedule_date) -> bool:
    for name, d, _reason in wb.availability:
        if name.strip().lower() == mechanic.strip().lower() and d == schedule_date:
            return True
    return False


def _assigned_partner(slot, slots) -> str:
    for s in slots:
        if (s.schedule_date == slot.schedule_date and s.location == slot.location
                and s.mechanic != slot.mechanic and s.mechanic):
            return s.mechanic
    return ""


def _consecutive_date_penalty(mechanic, schedule_date, slots) -> int:
    for s in slots:
        if s.mechanic == mechanic and abs((s.schedule_date - schedule_date).days) == 7:
            return 1
    return 0


def _pair_repeat_penalty(mechanic, slot, slots, pair_counts) -> int:
    partner = _assigned_partner(slot, slots)
    if partner:
        key = _pair_key(mechanic, partner)
        if key in pair_counts:
            return pair_counts[key]
    return 0


def _location_imbalance_after(mechanic, location, location_counts) -> int:
    one = location_counts.get(mechanic + "|" + LOCATION_ONE, 0)
    two = location_counts.get(mechanic + "|" + LOCATION_TWO, 0)
    if location == LOCATION_ONE:
        one += 1
    else:
        two += 1
    return abs(one - two)


def _count_spread_after(selected, total_counts) -> Tuple[int, int]:
    maximum = 0
    minimum = 2_147_483_647
    for mechanic, count in total_counts.items():
        c = count + 1 if mechanic == selected else count
        maximum = max(maximum, c)
        minimum = min(minimum, c)
    return maximum, minimum


def _days_since_last_assignment(last_date, schedule_date) -> int:
    return max(0, (schedule_date - last_date).days)


def _creates_same_day_duplicate(mechanic, changing_index, slots) -> bool:
    changing = slots[changing_index]
    for idx, s in enumerate(slots):
        if idx != changing_index and s.schedule_date == changing.schedule_date and s.mechanic == mechanic:
            return True
    return False


def _first_empty_row(wb):
    for row in wb.schedule:
        if not isinstance(row.get("date"), date):
            return row
    return None


def _date_key(value) -> str:
    if isinstance(value, date):
        return value.strftime("%Y%m%d")
    return str(value)


def _pair_key(a, b) -> str:
    return f"{a}|{b}" if a.lower() < b.lower() else f"{b}|{a}"


def _is_future(value, today) -> bool:
    return isinstance(value, date) and value >= today


def _blank(value) -> bool:
    return len(str(value or "").strip()) == 0


def _add_months(d: date, months: int) -> date:
    month = d.month - 1 + months
    year = d.year + month // 12
    month = month % 12 + 1
    day = min(d.day, [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                      31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1])
    return date(year, month, day)


# --- transparency: fairness summary + verification --------------------------


def fairness_summary(wb: Workbook) -> dict:
    """Statistical analysis of workload spread across active mechanics, computed
    live from the History sheet."""
    eligible = active_mechanics(wb)
    counts = _lifetime_counts(wb, eligible)
    values = list(counts.values())
    total = sum(values)
    if values:
        gap = max(values) - min(values)
        mean = total / len(values)
        stdev = (sum((v - mean) ** 2 for v in values) / len(values)) ** 0.5
    else:
        gap = 0
        stdev = 0.0
    return {
        "active_mechanics": len(eligible),
        "total_assignments": total,
        "gap": gap,          # busiest minus least-busy
        "stdev": round(stdev, 3),  # spread of workload (statistical analysis)
    }


def verify_generated_schedule(wb: Workbook, today: date) -> dict:
    """Post-generation work verification. Confirms:
      1. zero duplicate mechanics on the same date,
      2. zero unavailable/inactive mechanics scheduled,
      3. all future uncompleted rows filled.
    Returns a structured report plus a human-readable log.
    """
    active = set(active_mechanics(wb))
    duplicate_dates = []
    illegal = []          # (date, location, mechanic, reason)
    unfilled_rows = []

    for row in wb.schedule:
        d = row.get("date")
        if not _is_future(d, today):
            continue
        people_today = []
        for colname in (LOCATION_ONE, LOCATION_TWO):
            cell = str(row.get(colname) or "").strip()
            if cell == "":
                unfilled_rows.append((d, colname))
                continue
            names = [p.strip() for p in cell.split("/") if p.strip()]
            people_today.extend(names)
            for name in names:
                if name not in active:
                    illegal.append((d, colname, name, "inactive"))
                elif _is_unavailable(wb, name, d):
                    illegal.append((d, colname, name, "unavailable"))
        if len(people_today) != len(set(people_today)):
            duplicate_dates.append(d)

    report = {
        "duplicates_ok": len(duplicate_dates) == 0,
        "duplicate_dates": duplicate_dates,
        "no_illegal_ok": len(illegal) == 0,
        "illegal_assignments": illegal,
        "all_filled_ok": len(unfilled_rows) == 0,
        "unfilled_rows": unfilled_rows,
    }
    report["all_passed"] = (report["duplicates_ok"] and report["no_illegal_ok"]
                            and report["all_filled_ok"])
    report["log"] = (
        "Verification results:\n"
        f"  1. No duplicate mechanic on a date: {'PASS' if report['duplicates_ok'] else 'FAIL (' + str(len(duplicate_dates)) + ')'}\n"
        f"  2. No inactive/unavailable scheduled: {'PASS' if report['no_illegal_ok'] else 'FAIL (' + str(len(illegal)) + ')'}\n"
        f"  3. All future rows filled: {'PASS' if report['all_filled_ok'] else 'FAIL (' + str(len(unfilled_rows)) + ')'}"
    )
    return report


def generate_schedule(wb: Workbook, start_date: date, number_of_weeks: int, day_of_week: str,
                      today: date):
    """Full 'Generate Schedule' flow mirroring the VBA button: generate dates,
    check staffing, assign, verify, and summarise."""
    generate_dates(wb, start_date, number_of_weeks, day_of_week, today)
    check_active_staff_threshold(wb, today)
    slots = assign_future_schedule(wb, today)
    report = verify_generated_schedule(wb, today)
    summary = fairness_summary(wb)
    return slots, report, summary
