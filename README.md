# Assigned Late Schedule

A self-contained, macro-enabled Microsoft Excel workbook (`.xlsm`) that recreates
the paper **Assigned Late Schedule** (4:30 PM – 6:00 PM) and automatically fills
it in as fairly as mathematically possible.

The supervisor workflow is intentionally tiny:

1. Open the workbook.
2. Set **Start Date**, **Number of Weeks**, and **Day of Week** (Friday by default).
3. Click **Generate Schedule**.
4. Click **Print Schedule**.

Everything else — date generation, fair assignment, location/pair rotation, and
the permanent fairness history — happens automatically.

## Requirements

- **Microsoft Excel for Windows** (desktop `.xlsm` with macros enabled)
- **Not supported:** Excel for Mac — the scheduling engine uses Windows-only VBA
  features (`Scripting.Dictionary` and related COM). On Mac, the buttons will
  show a clear message instead of running.

## Worksheets

| Sheet | Purpose |
| --- | --- |
| **Schedule** | The printable paper-style schedule + the three settings and four buttons. Columns: `Date`, `1 Centre Street`, `100 Centre Street`. Each location shows two mechanics as `John Smith / Mike Jones`. |
| **Mechanics** | `Mechanic Name`, `Active (Yes/No)`. Add or deactivate mechanics here. |
| **Availability** | `Mechanic`, `Date`, `Reason`. Anyone listed for a date is never scheduled that date. |
| **History** *(hidden)* | `Date`, `Location`, `Mechanic`, `Assignment Created On`. Never edited by hand; it is the permanent record the fairness engine reads. |

## Buttons

- **Generate Schedule** – generate dates, assign every empty future slot, record history.
- **Generate Dates** – only insert the schedule dates.
- **Clear Future Schedule** – clear assignments for future dates only (history kept).
- **Print Schedule** – format the Schedule sheet to print on one page where practical.

## Fairness engine

Assignments are optimised, not rotated. A global pass assigns the whole horizon,
then balancing swaps run until no swap improves fairness. Priority order:
lowest lifetime assignments → longest time since last → lowest recent load →
avoid consecutive weeks → avoid repeat pairings → balance the two locations →
minimise the max/min assignment gap across all mechanics. History makes this
fair over months and years, not just within one schedule.

The full algorithm is in [`vba/AssignedLateSchedule.bas`](vba/AssignedLateSchedule.bas).

## Transparency, safeguards & verification

- **Fairness Summary block** – a panel beside the schedule (outside the print
  area) shows a statistical analysis of workload spread computed live from
  History: active mechanics, total assignments, the **assignment gap**
  (busiest minus least-busy) and the **workload standard deviation** (spread;
  lower is more even). Refreshed by `UpdateFairnessSummary` on each generate.
- **Audit logging** – `CommitAssignmentToHistory` writes a `Selection Note` on
  every History row explaining why each mechanic was chosen (lifetime/recent
  load, idle time, pairing rotation, location balance).
- **Duplicate prevention** – `WouldDuplicateOnDate` strictly blocks a mechanic
  from being placed at both buildings on the same date.
- **Active staff threshold** – `CheckActiveStaffThreshold` runs before assigning
  and aborts with a clear message if any generated date lacks enough active,
  available mechanics to fill its empty slots.
- **Post-generation verification** – `VerifyGeneratedSchedule` runs after
  optimisation and reports a validation log confirming: (1) zero duplicate
  mechanics on a date, (2) zero inactive/unavailable mechanics scheduled, and
  (3) all future rows filled. The log appears in the completion message.

## Building the workbook (developers)

Requires Python 3. No Microsoft Excel needed.

```bash
pip install -r requirements.txt

python tools/compile_vba.py   # compile vba/*.bas -> assets/vbaProject.bin
python build_workbook.py      # build "Assigned Late Schedule.xlsm"
```

- `build_workbook.py` creates the layout, tables, named ranges, data validation,
  print setup, and button→macro bindings, then embeds `assets/vbaProject.bin`.
- `tools/compile_vba.py` compiles the VBA source into `assets/vbaProject.bin`
  using the pure-Python `pyOpenVBA` library. Run it whenever the `.bas` changes.
  Excel recompiles the p-code from the embedded source on first open.

## Testing

The scheduling/fairness logic is mirrored in Python so it can be verified before
the VBA is compiled. Keep `tests/schedule_engine.py` in sync with the `.bas`.

```bash
python tests/test_schedule.py    # 16 fairness/behaviour tests
python tests/demo_schedule.py    # prints a sample generated schedule
```
