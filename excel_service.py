import logging
from datetime import date, datetime

from openpyxl import load_workbook

log = logging.getLogger(__name__)

LOCATION_SHEETS = ['Blanchardstown', 'Cork', 'Liffey Valley', 'Nutgrove', 'Whitewater']
INCOME_ROW      = 5   # Row where weekly income figures go


def _norm(name: str) -> str:
    return name.strip().lower()


def _find_week_columns(ws, target_sunday: date):
    """
    Scan row 3 for a datetime matching target_sunday.
    Returns (weekday_col, sunday_col) as 1-based indices, or (None, None).
    """
    for cell in ws[3]:
        if isinstance(cell.value, datetime) and cell.value.date() == target_sunday:
            sunday_col  = cell.column
            weekday_col = sunday_col - 1
            log.info(f"[{ws.title}] Sunday {target_sunday} → col {sunday_col}, Mon-Sat → col {weekday_col}")
            return weekday_col, sunday_col
    log.warning(f"[{ws.title}] No column found for Sunday {target_sunday}")
    return None, None


def _get_sheet_names(ws) -> list:
    names = []
    for row in ws.iter_rows(min_row=7, min_col=3, max_col=3):
        cell = row[0]
        if cell.value and isinstance(cell.value, str) and cell.value.strip():
            names.append(cell.value.strip())
    return names


def _find_staff_row(ws, name: str):
    target = _norm(name)
    for row in ws.iter_rows(min_row=7, min_col=3, max_col=3):
        cell = row[0]
        if cell.value and _norm(str(cell.value)) == target:
            return cell.row
    return None


def _match_name(square_name: str, sheet_names: list):
    sq_norm   = _norm(square_name)
    sq_parts  = sq_norm.split()
    sheet_map = {_norm(n): n for n in sheet_names}

    if sq_norm in sheet_map:
        return sheet_map[sq_norm]
    if sq_parts and sq_parts[0] in sheet_map:
        return sheet_map[sq_parts[0]]
    if len(sq_parts) > 1 and sq_parts[-1] in sheet_map:
        return sheet_map[sq_parts[-1]]
    return None


def _clear_week_columns(ws, weekday_col: int, sunday_col: int):
    """Wipe this week's two columns before writing so stale data never persists."""
    for row in ws.iter_rows(min_row=5, min_col=weekday_col, max_col=sunday_col):
        for cell in row:
            if cell.column in (weekday_col, sunday_col):
                if not isinstance(cell.value, str):
                    cell.value = None


def update_excel_wages(
    file_path: str,
    shifts: dict,
    income: dict,
    target_sunday: date = None,
) -> dict:
    """
    Write shift hours and weekly income into the wages Excel file.

    Args:
        file_path:     Path to the .xlsx wages file.
        shifts:        {sheet_name: {square_name: {weekday_hours, sunday_hours}}}
        income:        {sheet_name: income_float}  — Totals Collected - Gift Cards
        target_sunday: The Sunday of the week being processed.
    """
    if target_sunday is None:
        from square_service import get_previous_week_sunday
        target_sunday = get_previous_week_sunday()

    wb      = load_workbook(file_path)
    summary = {}

    for sheet_name in LOCATION_SHEETS:
        if sheet_name not in wb.sheetnames:
            log.warning(f"Sheet '{sheet_name}' missing — skipping")
            continue

        ws              = wb[sheet_name]
        location_shifts = shifts.get(sheet_name, {})
        location_income = income.get(sheet_name, 0.0)

        weekday_col, sunday_col = _find_week_columns(ws, target_sunday)
        if not weekday_col:
            summary[sheet_name] = {'error': f'Column not found for Sunday {target_sunday}'}
            continue

        # Wipe this week's columns first — prevents stale data
        _clear_week_columns(ws, weekday_col, sunday_col)

        # Write income into row 5 at the weekday column
        if location_income:
            ws.cell(row=INCOME_ROW, column=weekday_col).value = location_income
            log.info(f"[{sheet_name}] Income: €{location_income:.2f} → row {INCOME_ROW}, col {weekday_col}")

        if not location_shifts:
            summary[sheet_name] = {
                'updated': [], 'unmatched': [],
                'note': 'No shifts this week',
                'income': location_income,
            }
            continue

        sheet_names = _get_sheet_names(ws)
        updated   = []
        unmatched = []

        for sq_name, hours in location_shifts.items():
            matched = _match_name(sq_name, sheet_names)

            if not matched:
                unmatched.append(sq_name)
                log.warning(f"[{sheet_name}] Could not match Square name '{sq_name}'")
                continue

            row = _find_staff_row(ws, matched)
            if not row:
                unmatched.append(sq_name)
                continue

            weekday_h = round(hours['weekday_hours'], 2)
            sunday_h  = round(hours['sunday_hours'],  2)

            if weekday_h > 0:
                ws.cell(row=row, column=weekday_col).value = weekday_h
            if sunday_h > 0:
                ws.cell(row=row, column=sunday_col).value  = sunday_h

            updated.append({'name': matched, 'weekday_hours': weekday_h, 'sunday_hours': sunday_h})
            log.info(f"[{sheet_name}] {matched}: {weekday_h}h weekday, {sunday_h}h Sun")

        summary[sheet_name] = {
            'updated':   updated,
            'unmatched': unmatched,
            'income':    location_income,
        }
        log.info(f"[{sheet_name}] Done — {len(updated)} updated, {len(unmatched)} unmatched, income €{location_income:.2f}")

    wb.save(file_path)
    log.info(f"Saved workbook to {file_path}")
    return summary
