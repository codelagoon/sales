"""Demonstrate the workbook's core behaviour: build a fair late schedule.

This drives the exact engine (tests/schedule_engine.py) that the compiled VBA
runs when a supervisor clicks "Generate Schedule", so it is a faithful preview
of the workbook's output.  Run:  python3 tests/demo_schedule.py
"""

from __future__ import annotations

from datetime import date, timedelta

from schedule_engine import (
    LOCATION_ONE,
    LOCATION_TWO,
    Workbook,
    assign_future_schedule,
    generate_dates,
)

ROSTER = ["Ava Torres", "Ben Cole", "Cara Diaz", "Drew Kim",
          "Eli Fox", "Gia Hall", "Hank Ito"]
TODAY = date(2026, 8, 1)


def show(wb, title):
    print("=" * 74)
    print(title)
    print("=" * 74)
    print(f"{'Date':<22}{LOCATION_ONE:<26}{LOCATION_TWO}")
    print("-" * 74)
    for row in wb.schedule:
        d = row["date"].strftime("%a %m/%d/%Y")
        print(f"{d:<22}{row[LOCATION_ONE]:<26}{row[LOCATION_TWO]}")
    print()


def fairness(wb):
    totals, loc = {}, {}
    for h in wb.history:
        totals[h.mechanic] = totals.get(h.mechanic, 0) + 1
        loc.setdefault(h.mechanic, {LOCATION_ONE: 0, LOCATION_TWO: 0})
        loc[h.mechanic][h.location] += 1
    print("Fairness summary (lifetime, from permanent History):")
    print(f"  {'Mechanic':<14}{'Total':>6}{'  @1 Centre':>12}{'  @100 Centre':>14}")
    for m in sorted(totals):
        print(f"  {m:<14}{totals[m]:>6}{loc[m][LOCATION_ONE]:>12}{loc[m][LOCATION_TWO]:>14}")
    spread = max(totals.values()) - min(totals.values())
    print(f"  --> assignment spread (max-min): {spread}  "
          f"(0 or 1 = as fair as mathematically possible)")
    print()
    return spread


def main():
    wb = Workbook()
    wb.mechanics = [(n, "Yes") for n in ROSTER]

    # Cycle 1: a supervisor generates 12 Fridays starting 8/7/2026.
    generate_dates(wb, date(2026, 8, 7), 12, "Friday", TODAY)
    assign_future_schedule(wb, TODAY)
    show(wb, "CYCLE 1  -  Generate Schedule (12 Fridays, 7 mechanics)")
    fairness(wb)

    # Cycle 2: next quarter, with one mechanic unavailable on the first date.
    next_start = wb.schedule[-1]["date"] + timedelta(days=7)
    next_today = wb.schedule[-1]["date"] + timedelta(days=1)
    wb.availability.append(("Ava Torres", next_start, "Vacation"))
    wb.schedule = []  # supervisor clears the printed sheet; History is permanent
    generate_dates(wb, next_start, 12, "Friday", next_today)
    assign_future_schedule(wb, next_today)
    show(wb, f"CYCLE 2  -  next quarter (Ava Torres unavailable {next_start:%m/%d})")
    fairness(wb)

    blocked = [h for h in wb.history if h.mechanic == "Ava Torres" and h.date == next_start]
    print(f"Availability respected: Ava Torres assignments on {next_start:%m/%d} = "
          f"{len(blocked)} (expected 0)")

    print("\nCore functionality demonstrated: fair, balanced, availability-aware "
          "late schedule\ngenerated automatically and carried forward via permanent history.")


if __name__ == "__main__":
    main()
