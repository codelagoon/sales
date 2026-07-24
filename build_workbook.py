"""Build the Assigned Late Schedule macro-enabled Excel workbook.

The VBA source lives in vba/AssignedLateSchedule.bas.  XlsxWriter cannot
compile VBA, so the module is pre-compiled into `assets/vbaProject.bin` by
`tools/compile_vba.py` (pure Python, no Excel required); run that script after
changing the VBA.  The layout, tables, names, validation, print setup, and
macro button bindings are created here and the compiled project is embedded.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import xlsxwriter


ROOT = Path(__file__).parent
OUTPUT = ROOT / "Assigned Late Schedule.xlsm"
VBA_PROJECT = ROOT / "assets" / "vbaProject.bin"


def build_workbook() -> None:
    if not VBA_PROJECT.exists():
        raise FileNotFoundError(
            "assets/vbaProject.bin is required to create a macro-enabled workbook."
        )

    workbook = xlsxwriter.Workbook(str(OUTPUT))
    workbook.add_vba_project(str(VBA_PROJECT))
    workbook.set_properties(
        {
            "title": "Assigned Late Schedule",
            "subject": "Fair late-shift schedule",
            "author": "Operations",
            "comments": "Automated late schedule with fairness history.",
        }
    )

    _add_schedule_sheet(workbook)
    _add_mechanics_sheet(workbook)
    _add_availability_sheet(workbook)
    _add_history_sheet(workbook)
    workbook.close()


def _add_schedule_sheet(workbook: xlsxwriter.Workbook) -> None:
    sheet = workbook.add_worksheet("Schedule")
    sheet.set_tab_color("#1F4E78")
    sheet.hide_gridlines(2)
    sheet.set_landscape()
    sheet.fit_to_pages(1, 1)
    sheet.set_paper(9)  # A4
    sheet.set_margins(0.35, 0.35, 0.5, 0.5)
    sheet.repeat_rows(0, 9)
    sheet.set_header("&C&\"Arial,Bold\"Assigned Late Schedule")
    sheet.set_footer("&LGenerated &D&RPage &P of &N")
    sheet.set_column("A:A", 16)
    sheet.set_column("B:C", 37)
    sheet.set_column("D:D", 2)
    sheet.set_column("E:E", 20)

    navy = "#1F4E78"
    border = "#7F7F7F"
    title = workbook.add_format(
        {
            "bold": True,
            "font_size": 18,
            "font_color": "#FFFFFF",
            "bg_color": navy,
            "align": "center",
            "valign": "vcenter",
        }
    )
    subtitle = workbook.add_format(
        {
            "bold": True,
            "font_size": 11,
            "font_color": "#FFFFFF",
            "bg_color": navy,
            "align": "center",
            "valign": "vcenter",
        }
    )
    label = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": navy,
            "border": 1,
            "border_color": border,
            "align": "right",
        }
    )
    input_format = workbook.add_format(
        {
            "bg_color": "#FFF2CC",
            "border": 1,
            "border_color": border,
            "align": "left",
            "num_format": "m/d/yyyy",
        }
    )
    input_number = workbook.add_format(
        {"bg_color": "#FFF2CC", "border": 1, "border_color": border, "align": "left"}
    )
    note = workbook.add_format({"font_color": "#595959", "italic": True, "font_size": 9})
    instruction = workbook.add_format(
        {"font_color": "#1F1F1F", "font_size": 9, "text_wrap": True}
    )
    table_header = workbook.add_format(
        {
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": navy,
            "border": 1,
            "border_color": border,
            "align": "center",
            "valign": "vcenter",
        }
    )
    table_date = workbook.add_format(
        {
            "border": 1,
            "border_color": border,
            "align": "center",
            "valign": "vcenter",
            "num_format": "dddd, m/d/yyyy",
        }
    )
    table_assignment = workbook.add_format(
        {
            "border": 1,
            "border_color": border,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        }
    )

    sheet.merge_range("A1:C1", "Assigned Late Schedule", title)
    sheet.merge_range("A2:C2", "4:30 PM - 6:00 PM", subtitle)
    sheet.set_row(0, 26)
    sheet.set_row(1, 20)
    sheet.write("A4", "Start Date", label)
    sheet.write_datetime("B4", date.today(), input_format)
    sheet.write("A5", "Number of Weeks", label)
    sheet.write_number("B5", 12, input_number)
    sheet.write("A6", "Day of Week", label)
    sheet.write("B6", "Friday", input_number)
    sheet.write("A7", "Instructions", label)
    sheet.merge_range(
        "B7:C7",
        "Set the three yellow fields, then click Generate Schedule. "
        "Use Print Schedule when ready.",
        instruction,
    )
    sheet.data_validation(
        "B6",
        {
            "validate": "list",
            "source": [
                "Monday",
                "Tuesday",
                "Wednesday",
                "Thursday",
                "Friday",
                "Saturday",
                "Sunday",
            ],
        },
    )
    sheet.data_validation("B5", {"validate": "integer", "criteria": "between", "minimum": 1, "maximum": 104})
    workbook.define_name("ScheduleStartDate", "=Schedule!$B$4")
    workbook.define_name("ScheduleWeekCount", "=Schedule!$B$5")
    workbook.define_name("ScheduleDayOfWeek", "=Schedule!$B$6")
    sheet.write("A9", "Future dates and assignments appear below.", note)

    sheet.add_table(
        "A11:C12",
        {
            "name": "ScheduleTable",
            "style": "Table Style Medium 2",
            "columns": [
                {"header": "Date", "header_format": table_header, "format": table_date},
                {
                    "header": "1 Centre Street",
                    "header_format": table_header,
                    "format": table_assignment,
                },
                {
                    "header": "100 Centre Street",
                    "header_format": table_header,
                    "format": table_assignment,
                },
            ],
        },
    )
    sheet.set_row(10, 24)
    sheet.set_row(11, 28)

    sheet.insert_button(
        "E4",
        {"macro": "GenerateSchedule", "caption": "Generate Schedule", "width": 150, "height": 24, "font": {"bold": True, "color": "#FFFFFF"}, "fill": {"color": "#4472C4"}, "line": {"color": "#1F4E78"}},
    )
    sheet.insert_button(
        "E5",
        {"macro": "GenerateDates", "caption": "Generate Dates", "width": 150, "height": 24, "font": {"bold": True, "color": "#FFFFFF"}, "fill": {"color": "#4472C4"}, "line": {"color": "#1F4E78"}},
    )
    sheet.insert_button(
        "E6",
        {"macro": "ClearFutureSchedule", "caption": "Clear Future Schedule", "width": 150, "height": 24, "font": {"bold": True, "color": "#FFFFFF"}, "fill": {"color": "#4472C4"}, "line": {"color": "#1F4E78"}},
    )
    sheet.insert_button(
        "E7",
        {"macro": "PrintSchedule", "caption": "Print Schedule", "width": 150, "height": 24, "font": {"bold": True, "color": "#FFFFFF"}, "fill": {"color": "#4472C4"}, "line": {"color": "#1F4E78"}},
    )
    sheet.print_area("A1:C12")
    sheet.freeze_panes(10, 0)


def _add_mechanics_sheet(workbook: xlsxwriter.Workbook) -> None:
    sheet = workbook.add_worksheet("Mechanics")
    sheet.set_tab_color("#70AD47")
    sheet.freeze_panes(1, 0)
    sheet.set_column("A:A", 34)
    sheet.set_column("B:B", 16)
    sheet.add_table(
        "A1:B2",
        {
            "name": "MechanicsTable",
            "style": "Table Style Medium 4",
            "columns": [{"header": "Mechanic Name"}, {"header": "Active (Yes/No)"}],
        },
    )
    sheet.data_validation(
        "B2:B1048576", {"validate": "list", "source": ["Yes", "No"], "ignore_blank": True}
    )


def _add_availability_sheet(workbook: xlsxwriter.Workbook) -> None:
    sheet = workbook.add_worksheet("Availability")
    sheet.set_tab_color("#ED7D31")
    sheet.freeze_panes(1, 0)
    sheet.set_column("A:A", 34)
    sheet.set_column("B:B", 16)
    sheet.set_column("C:C", 48)
    date_format = workbook.add_format({"num_format": "m/d/yyyy"})
    sheet.add_table(
        "A1:C2",
        {
            "name": "AvailabilityTable",
            "style": "Table Style Medium 6",
            "columns": [
                {"header": "Mechanic"},
                {"header": "Date", "format": date_format},
                {"header": "Reason"},
            ],
        },
    )


def _add_history_sheet(workbook: xlsxwriter.Workbook) -> None:
    sheet = workbook.add_worksheet("History")
    sheet.set_tab_color("#A5A5A5")
    sheet.set_column("A:A", 16)
    sheet.set_column("B:B", 22)
    sheet.set_column("C:C", 34)
    sheet.set_column("D:D", 22)
    date_format = workbook.add_format({"num_format": "m/d/yyyy h:mm AM/PM"})
    sheet.add_table(
        "A1:D2",
        {
            "name": "HistoryTable",
            "style": "Table Style Medium 15",
            "columns": [
                {"header": "Date", "format": date_format},
                {"header": "Location"},
                {"header": "Mechanic"},
                {"header": "Assignment Created On", "format": date_format},
            ],
        },
    )
    sheet.hide()


if __name__ == "__main__":
    build_workbook()
