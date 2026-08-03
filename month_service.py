"""
month_service.py

Each month's file contains ONLY that month's working days.
When a week crosses a month boundary the data is split across two files:
  - previous month's file gets the Mon→last_day_of_prev_month portion
  - current month's file gets the first_day_of_curr_month→Sat portion + Sunday

Column rule:
  - Partial START = from first day of THIS month to Saturday of that week
    (no previous-month days — those live in the previous month's file)
    Exception: months starting on Monday show previous Fri-Sat + Sun (payroll convention)
  - Partial END   = from Monday of the last week to the last day of THIS month
    (no next-month days — those live in the next month's file)
"""

import calendar
import logging
import shutil
from datetime import date, timedelta, datetime
from pathlib import Path

from openpyxl import load_workbook

log = logging.getLogger(__name__)

LOCATION_SHEETS = ['Blanchardstown', 'Cork', 'Liffey Valley', 'Nutgrove', 'Whitewater']
MONTH_ROW       = 1
MONTH_COL       = 6
PERIOD_ROW      = 2
DATE_ROW        = 3
STAFF_ROW_START = 7
STAFF_ROW_END   = 25
DATA_COL_START  = 6    # F
DATA_COL_END    = 16   # P  (max 11 data columns)

MONTH_NAMES = ['', 'January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
MONTH_ABBR  = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
DAY_ABBR    = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


def _range_label(start: date, end: date) -> str:
    if start.month == end.month:
        return f"{start.day:02d}-{end.day:02d} {MONTH_ABBR[start.month]}"
    return f"{start.day:02d} {MONTH_ABBR[start.month]}-{end.day:02d} {MONTH_ABBR[end.month]}"


def _dow_label(start: date, end: date) -> str:
    s = DAY_ABBR[start.weekday()]
    e = DAY_ABBR[end.weekday()]
    return f"{s}-{e}" if s != e else s


def generate_week_columns(year: int, month: int) -> list:
    """
    Build column definitions for a given month.

    Partial start: from the first day of the month to the Saturday of that week.
    Partial end:   from Monday of the last week to the last day of the month.
    No days from adjacent months appear — cross-month data lives in those months' files.

    Exception: months starting on Monday use the payroll convention of showing
    the previous week's Fri-Sat + Sun as the first columns.

    Returns list of (row2_label, row3_value, col_type):
      col_type: 'weekday' | 'sunday'
    """
    first_day = date(year, month, 1)
    last_day  = date(year, month, calendar.monthrange(year, month)[1])
    dow       = first_day.weekday()   # Mon=0

    if dow == 0:
        # Month starts Monday — payroll convention: show prev Fri-Sat + Sun
        p_start = first_day - timedelta(days=3)   # Friday (prev month)
        p_end   = first_day - timedelta(days=2)   # Saturday (prev month)
        sunday  = first_day - timedelta(days=1)   # Sunday (prev month)
    else:
        # Month starts Tue-Sun:
        # Partial start = first_day to Saturday of that week (all in this month)
        p_start = first_day
        p_end   = first_day + timedelta(days=(5 - dow))   # Saturday of that week
        sunday  = p_end + timedelta(days=1)                # Sunday of that week

    cols = [
        (_dow_label(p_start, p_end), _range_label(p_start, p_end), 'weekday'),
        ('Sun', datetime(sunday.year, sunday.month, sunday.day), 'sunday'),
    ]

    current_monday = sunday + timedelta(days=1)

    while current_monday <= last_day:
        sat = current_monday + timedelta(days=5)
        sun = current_monday + timedelta(days=6)

        if sat > last_day:
            # Partial end: Mon to last day of THIS month only
            pe_end = last_day
            cols.append((_dow_label(current_monday, pe_end),
                         _range_label(current_monday, pe_end),
                         'weekday'))
            break
        else:
            cols.append(('Mon-Sat', _range_label(current_monday, sat), 'weekday'))
            cols.append(('Sun', datetime(sun.year, sun.month, sun.day), 'sunday'))

        current_monday += timedelta(days=7)

    return cols


def _refresh_location_sheet(ws, year: int, month: int):
    ws.cell(row=MONTH_ROW, column=MONTH_COL).value = MONTH_NAMES[month].upper()

    cols = generate_week_columns(year, month)

    for col in range(DATA_COL_START, DATA_COL_END + 1):
        ws.cell(row=PERIOD_ROW, column=col).value = None
        ws.cell(row=DATE_ROW,   column=col).value = None

    for i, (label, val, _) in enumerate(cols):
        col = DATA_COL_START + i
        ws.cell(row=PERIOD_ROW, column=col).value = label
        ws.cell(row=DATE_ROW,   column=col).value = val

    for row in range(STAFF_ROW_START, STAFF_ROW_END + 1):
        for col in range(DATA_COL_START, DATA_COL_END + 1):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell.value, str):
                cell.value = None

    log.info(f"[{ws.title}] Refreshed for {MONTH_NAMES[month]} {year} ({len(cols)} cols)")


def refresh_for_month(master_path: str, output_path: str, year: int, month: int):
    shutil.copy2(master_path, output_path)
    wb = load_workbook(output_path)

    for sheet_name in LOCATION_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        _refresh_location_sheet(wb[sheet_name], year, month)

    if 'Final Payment' in wb.sheetnames:
        dow = date(year, month, 1).weekday()
        ps  = (date(year, month, 1) - timedelta(days=3) if dow == 0
               else date(year, month, 1))
        wb['Final Payment'].cell(row=1, column=2).value = datetime(
            ps.year, ps.month, ps.day)

    wb.save(output_path)
    log.info(f"Generated {Path(output_path).name}")


def monthly_filename(year: int, month: int) -> str:
    return f"MB_Ireland_{MONTH_NAMES[month][:3]}_{year}.xlsx"


def current_expected_filename() -> str:
    today = datetime.now()
    return monthly_filename(today.year, today.month)
