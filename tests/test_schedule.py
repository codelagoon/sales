"""Comprehensive tests for the Assigned Late Schedule fairness engine.

Run directly:  python3 tests/test_schedule.py
Or with pytest: pytest tests/test_schedule.py

These tests validate the algorithm described in the workbook spec BEFORE the
VBA is compiled into vbaProject.bin.  They mirror the exact scoring used by the
macro (see tests/schedule_engine.py).
"""

from __future__ import annotations

import sys
from datetime import date, timedelta

from schedule_engine import (
    LOCATION_ONE,
    LOCATION_TWO,
    Workbook,
    HistoryRow,
    assign_future_schedule,
    first_matching_date,
    generate_dates,
    weekday_number,
)

TODAY = date(2026, 1, 1)  # fixed "today" so tests are deterministic


def make_workbook(names, active=None):
    wb = Workbook()
    active = active or {n: "Yes" for n in names}
    wb.mechanics = [(n, active.get(n, "Yes")) for n in names]
    return wb


def counts_by_mechanic(wb):
    result = {}
    for h in wb.history:
        result[h.mechanic] = result.get(h.mechanic, 0) + 1
    return result


def location_counts(wb):
    result = {}
    for h in wb.history:
        result.setdefault(h.mechanic, {LOCATION_ONE: 0, LOCATION_TWO: 0})
        result[h.mechanic][h.location] += 1
    return result


def pair_counts(wb):
    grouped = {}
    for h in wb.history:
        grouped.setdefault((h.date, h.location), []).append(h.mechanic)
    pairs = {}
    for people in grouped.values():
        if len(people) == 2:
            key = tuple(sorted(people))
            pairs[key] = pairs.get(key, 0) + 1
    return pairs


# --- tests -------------------------------------------------------------------


def test_dates_generated_on_correct_weekday():
    wb = make_workbook(["A", "B", "C", "D"])
    generate_dates(wb, start_date=date(2026, 8, 7), number_of_weeks=12,
                   day_of_week="Friday", today=TODAY)
    dates = [r["date"] for r in wb.schedule]
    assert len(dates) == 12, f"expected 12 dates, got {len(dates)}"
    assert all(d.weekday() == 4 for d in dates), "all dates must be Fridays"
    # spacing is 7 days and sorted ascending
    for earlier, later in zip(dates, dates[1:]):
        assert (later - earlier).days == 7
    assert dates[0] == date(2026, 8, 7), dates[0]


def test_dates_respect_other_weekdays():
    for day, py_wd in [("Monday", 0), ("Wednesday", 2), ("Sunday", 6)]:
        wb = make_workbook(["A", "B", "C", "D"])
        generate_dates(wb, start_date=date(2026, 8, 3), number_of_weeks=6,
                       day_of_week=day, today=TODAY)
        dates = [r["date"] for r in wb.schedule]
        assert len(dates) == 6
        assert all(d.weekday() == py_wd for d in dates), f"{day} generated wrong weekday"


def test_first_matching_date_is_on_or_after_start():
    # start on a Wednesday, ask for Friday -> two days later
    assert first_matching_date(date(2026, 8, 5), weekday_number("Friday")) == date(2026, 8, 7)
    # start already on the requested day -> same day
    assert first_matching_date(date(2026, 8, 7), weekday_number("Friday")) == date(2026, 8, 7)


def test_full_day_needs_four_distinct_mechanics():
    wb = make_workbook(["A", "B", "C"])
    generate_dates(wb, date(2026, 8, 7), 4, "Friday", TODAY)
    raised = False
    try:
        assign_future_schedule(wb, TODAY)
    except ValueError as exc:
        raised = "four active mechanics" in str(exc)
    assert raised, "should refuse to build with fewer than four active mechanics"


def test_no_mechanic_twice_on_same_date():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    by_date = {}
    for h in wb.history:
        by_date.setdefault(h.date, []).append(h.mechanic)
    for d, people in by_date.items():
        assert len(people) == len(set(people)), f"duplicate mechanic on {d}: {people}"
        assert len(people) == 4, f"expected 4 assignments on {d}, got {len(people)}"


def test_two_names_per_location_cell():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 8, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    for row in wb.schedule:
        for col in (LOCATION_ONE, LOCATION_TWO):
            assert row[col].count("/") == 1, f"cell must contain two names: {row[col]!r}"
            a, b = [p.strip() for p in row[col].split("/")]
            assert a and b and a != b


def test_lifetime_totals_are_balanced_when_divisible():
    # 6 mechanics, 12 weeks -> 48 slots / 6 = exactly 8 each -> spread 0.
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    counts = counts_by_mechanic(wb)
    assert len(counts) == 6
    spread = max(counts.values()) - min(counts.values())
    assert spread == 0, f"expected perfectly even totals, got {counts}"


def test_lifetime_totals_within_one_when_not_divisible():
    # 5 mechanics, 12 weeks -> 48 slots / 5 = 9.6 -> spread must be <= 1.
    wb = make_workbook(["A", "B", "C", "D", "E"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    counts = counts_by_mechanic(wb)
    spread = max(counts.values()) - min(counts.values())
    assert spread <= 1, f"expected spread <= 1, got {counts}"


def test_unavailable_mechanic_is_never_scheduled():
    wb = make_workbook(["A", "B", "C", "D", "E"])
    generate_dates(wb, date(2026, 8, 7), 10, "Friday", TODAY)
    blocked_dates = [r["date"] for r in wb.schedule[:3]]
    for d in blocked_dates:
        wb.availability.append(("A", d, "Vacation"))
    assign_future_schedule(wb, TODAY)
    for h in wb.history:
        if h.mechanic == "A":
            assert h.date not in blocked_dates, f"A scheduled on blocked date {h.date}"


def test_location_balance_per_mechanic():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    for mech, locs in location_counts(wb).items():
        gap = abs(locs[LOCATION_ONE] - locs[LOCATION_TWO])
        assert gap <= 2, f"{mech} location imbalance too high: {locs}"


def test_pairings_rotate():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    pairs = pair_counts(wb)
    # With 6 mechanics over 12 weeks no single pairing should dominate.
    assert max(pairs.values()) <= 4, f"a pairing repeated too often: {pairs}"


def test_completed_rows_never_overwritten():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    # Pre-fill the first future row as if it were already completed manually.
    wb.schedule[0][LOCATION_ONE] = "Zeb / Yan"
    wb.schedule[0][LOCATION_TWO] = "Xia / Wes"
    frozen_date = wb.schedule[0]["date"]
    assign_future_schedule(wb, TODAY)
    assert wb.schedule[0][LOCATION_ONE] == "Zeb / Yan"
    assert wb.schedule[0][LOCATION_TWO] == "Xia / Wes"
    # None of the pre-filled names are active mechanics, so history for that
    # date must not contain engine-made assignments.
    for h in wb.history:
        assert not (h.date == frozen_date), "completed row must not be re-added to history"


def test_history_grows_and_dedupes_on_rerun():
    wb = make_workbook(["A", "B", "C", "D", "E", "F"])
    generate_dates(wb, date(2026, 8, 7), 6, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    first = len(wb.history)
    assert first == 24, f"6 weeks x 4 = 24 history rows expected, got {first}"
    # Re-running with everything already filled adds nothing (no empty slots).
    assign_future_schedule(wb, TODAY)
    assert len(wb.history) == first, "re-run must not duplicate history"


def test_long_term_fairness_over_many_cycles():
    """Simulate a year of monthly generations feeding history forward."""
    names = ["A", "B", "C", "D", "E"]  # deliberately not a divisor of 4/week
    wb = make_workbook(names)
    today = date(2026, 1, 1)
    cycle_start = date(2026, 2, 6)  # a Friday
    for _ in range(12):  # 12 monthly cycles, 4 weeks each
        generate_dates(wb, cycle_start, 4, "Friday", today)
        assign_future_schedule(wb, today)
        # advance simulated clock and next start beyond the generated dates
        last_date = max(r["date"] for r in wb.schedule)
        today = last_date + timedelta(days=1)
        cycle_start = last_date + timedelta(days=7)
        # clear the schedule sheet for the next cycle (history is permanent)
        wb.schedule = []
    counts = counts_by_mechanic(wb)
    total = sum(counts.values())
    assert total == 12 * 4 * 4, f"unexpected total assignments {total}"
    spread = max(counts.values()) - min(counts.values())
    # Over a long horizon the permanent history should keep everyone within 1.
    assert spread <= 1, f"long-term spread should stay <=1, got {counts} (spread {spread})"


def test_history_biases_future_toward_underloaded():
    """A mechanic with heavy prior history should be de-prioritised."""
    wb = make_workbook(["A", "B", "C", "D", "E"])
    # Give A a big head start in history.
    past = date(2025, 6, 6)
    for i in range(8):
        wb.history.append(HistoryRow(past + timedelta(days=7 * i), LOCATION_ONE, "A", "seed"))
    generate_dates(wb, date(2026, 8, 7), 4, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    # A already had 8; others start at 0. Engine should give A the fewest new ones.
    a_new = sum(1 for h in wb.history if h.mechanic == "A" and h.date >= date(2026, 8, 7))
    others_new = {m: sum(1 for h in wb.history if h.mechanic == m and h.date >= date(2026, 8, 7))
                  for m in ["B", "C", "D", "E"]}
    assert a_new <= min(others_new.values()), (
        f"overloaded mechanic A got too many new assignments: A={a_new}, others={others_new}")


def test_determinism():
    def run():
        wb = make_workbook(["A", "B", "C", "D", "E", "F"])
        generate_dates(wb, date(2026, 8, 7), 10, "Friday", TODAY)
        assign_future_schedule(wb, TODAY)
        return [(r["date"], r[LOCATION_ONE], r[LOCATION_TWO]) for r in wb.schedule]
    assert run() == run(), "engine must be deterministic for identical inputs"


def main():
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_") and callable(v)]
    failures = 0
    for t in tests:
        try:
            t()
            print(f"PASS  {t.__name__}")
        except AssertionError as exc:
            failures += 1
            print(f"FAIL  {t.__name__}: {exc}")
        except Exception as exc:  # noqa: BLE001
            failures += 1
            print(f"ERROR {t.__name__}: {type(exc).__name__}: {exc}")
    print(f"\n{len(tests) - failures}/{len(tests)} tests passed")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
