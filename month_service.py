"""
month_service.py

Handles auto-generation of monthly wages files from the master template.
Clears all hour data and regenerates date headers for any given month.
"""

import calendar
import logging
import shutil
from datetime import date, timedelta, datetime
from pathlib import Path

from openpyxl import load_workbook

log = logging.getLogger(__name__)

# ── Constants (matching the Excel template structure) ─────────────────────────
LOCATION_SHEETS  = ['Blanchardstown', 'Cork', 'Liffey Valley', 'Nutgrove', 'Whitewater']
MONTH_ROW        = 1    # Row with the month name (e.g., "JUNE")
MONTH_COL        = 6    # Column F
PERIOD_ROW       = 2    # Row with period labels ("Mon-Sat", "Sun", etc.)
DATE_ROW         = 3    # Row with actual date values / date ranges
STAFF_ROW_START  = 7    # First staff data row
STAFF_ROW_END    = 25   # Last possible staff row (covers all locations)
DATA_COL_START   = 6    # Column F  — first hour-data column
DATA_COL_END     = 16   # Column P  — last  hour-data column (max for any month)

MONTH_NAMES = ['', 'January', 'February', 'March', 'April', 'May', 'June',
               'July', 'August', 'September', 'October', 'November', 'December']
MONTH_ABBR  = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
               'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
DAY_ABBR    = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']


# ── Date helpers ──────────────────────────────────────────────────────────────

def _range_label(start: date, end: date) -> str:
    if start.month == end.month:
        return f"{start.day:02d}-{end.day:02d} {MONTH_ABBR[start.month]}"
    return f"{start.day:02d} {MONTH_ABBR[start.month]}-{end.day:02d} {MONTH_ABBR[end.month]}"


def generate_week_columns(year: int, month: int) -> list:
    """
    Build the column structure (row 2 label + row 3 value) for a given month.

    Returns a list of (period_label, date_value) where:
      - period_label (str):  e.g., 'Fri - Sat', 'Mon-Sat', 'Sun', 'Mon-Tue'
      - date_value:          str (date range) for weekday cols, datetime for Sunday cols

    The list starts at column F and runs left-to-right.
    Max 11 entries (covers months with 5 full weeks + partial start/end).
    """
    first_day = date(year, month, 1)
    last_day  = date(year, month, calendar.monthrange(year, month)[1])
    dow       = first_day.weekday()  # Mon=0, Sun=6

    # ── Partial start + first Sunday ──────────────────────────────────────────
    if dow == 0:
        # Month starts Monday → capture previous Fri-Sat + Sunday
        partial_start = first_day - timedelta(days=3)   # Friday
        partial_end   = first_day - timedelta(days=2)   # Saturday
        sunday        = first_day - timedelta(days=1)   # Sunday
    else:
        # Month starts Tue–Sun → capture Mon-to-day_before_1st + Sunday of that week
        partial_start = first_day - timedelta(days=dow)       # Monday
        partial_end   = first_day - timedelta(days=1)          # Day before 1st
        sunday        = partial_start + timedelta(days=6)      # Sunday of that week

    s = DAY_ABBR[partial_start.weekday()]
    e = DAY_ABBR[partial_end.weekday()]
    partial_label = f"{s} - {e}" if s != e else s

    cols = [
        (partial_label, _range_label(partial_start, partial_end)),
        ('Sun', datetime(sunday.year, sunday.month, sunday.day)),
    ]

    # ── Full weeks + partial end ───────────────────────────────────────────────
    current_monday = sunday + timedelta(days=1)

    while current_monday <= last_day:
        sat = current_monday + timedelta(days=5)
        sun = current_monday + timedelta(days=6)

        if sat > last_day:
            # Partial end — last days of the month
            pe_end = last_day
            ps = DAY_ABBR[current_monday.weekday()]
            pe = DAY_ABBR[pe_end.weekday()]
            label = f"{ps}-{pe}" if ps != pe else ps
            cols.append((label, _range_label(current_monday, pe_end)))
            break
        else:
            cols.append(('Mon-Sat', _range_label(current_monday, sat)))
            cols.append(('Sun', datetime(sun.year, sun.month, sun.day)))

        current_monday += timedelta(days=7)

    return cols


# ── Core refresh logic ────────────────────────────────────────────────────────

def _refresh_location_sheet(ws, year: int, month: int):
    """Clear hours and update date headers on a single location sheet."""
    # 1. Update month name (row 1, col F)
    ws.cell(row=MONTH_ROW, column=MONTH_COL).value = MONTH_NAMES[month].upper()

    # 2. Generate new column definitions
    cols = generate_week_columns(year, month)

    # 3. Clear ALL cells in rows 2–3 between col F and col P
    for col in range(DATA_COL_START, DATA_COL_END + 1):
        ws.cell(row=PERIOD_ROW, column=col).value = None
        ws.cell(row=DATE_ROW,   column=col).value = None

    # 4. Write new period headers (row 2) and date values (row 3)
    for i, (label, val) in enumerate(cols):
        col = DATA_COL_START + i
        ws.cell(row=PERIOD_ROW, column=col).value = label
        ws.cell(row=DATE_ROW,   column=col).value = val

    # 5. Clear all hour values from staff rows (col F–P only, preserve formulas elsewhere)
    for row in range(STAFF_ROW_START, STAFF_ROW_END + 1):
        for col in range(DATA_COL_START, DATA_COL_END + 1):
            cell = ws.cell(row=row, column=col)
            if not isinstance(cell.value, str):  # don't wipe formula strings
                cell.value = None

    log.info(f"[{ws.title}] Refreshed for {MONTH_NAMES[month]} {year} ({len(cols)} data columns)")


def refresh_for_month(master_path: str, output_path: str, year: int, month: int):
    """
    Create a fresh monthly wages file from the master template.

    Args:
        master_path:  Path to master_template.xlsx (never modified).
        output_path:  Path to save the new monthly file.
        year, month:  Target month.
    """
    shutil.copy2(master_path, output_path)
    wb = load_workbook(output_path)

    for sheet_name in LOCATION_SHEETS:
        if sheet_name not in wb.sheetnames:
            log.warning(f"Sheet '{sheet_name}' not in master template — skipping")
            continue
        _refresh_location_sheet(wb[sheet_name], year, month)

    # Update the payroll-period start date in Final Payment sheet (row 1, col B)
    if 'Final Payment' in wb.sheetnames:
        cols = generate_week_columns(year, month)
        period_start_str = cols[0][1]  # date range string of the partial start
        # Store the actual date of the partial start period start
        dow = date(year, month, 1).weekday()
        partial_start_date = (
            date(year, month, 1) - timedelta(days=3) if dow == 0
            else date(year, month, 1) - timedelta(days=dow)
        )
        wb['Final Payment'].cell(row=1, column=2).value = datetime(
            partial_start_date.year, partial_start_date.month, partial_start_date.day
        )

    wb.save(output_path)
    log.info(f"Generated {Path(output_path).name}")


# ── Filename helpers ──────────────────────────────────────────────────────────

def monthly_filename(year: int, month: int) -> str:
    return f"MB_Ireland_{MONTH_NAMES[month][:3]}_{year}.xlsx"


def current_expected_filename() -> str:
    today = datetime.now()
    return monthly_filename(today.year, today.month)
