import os
import logging
from datetime import datetime, timedelta, date

import pytz
import requests

log = logging.getLogger(__name__)

MELBOURNE_TZ = pytz.timezone('Australia/Melbourne')
IRELAND_TZ   = pytz.timezone('Europe/Dublin')
SQUARE_BASE  = 'https://connect.squareup.com/v2'

LOCATION_MAP = {
    'Blanchardstown': os.getenv('SQUARE_LOC_BLANCHARDSTOWN'),
    'Cork':           os.getenv('SQUARE_LOC_CORK'),
    'Liffey Valley':  os.getenv('SQUARE_LOC_LIFFEY_VALLEY'),
    'Nutgrove':       os.getenv('SQUARE_LOC_NUTGROVE'),
    'Whitewater':     os.getenv('SQUARE_LOC_WHITEWATER'),
}


def _headers():
    return {
        'Authorization': f'Bearer {os.getenv("SQUARE_ACCESS_TOKEN")}',
        'Content-Type': 'application/json',
        'Square-Version': '2024-10-17',
    }


def get_previous_week_sunday() -> date:
    today = datetime.now(MELBOURNE_TZ).date()
    return today - timedelta(days=1)


def _get_week_range(target_sunday=None):
    prev_sunday = target_sunday or get_previous_week_sunday()
    prev_monday = prev_sunday - timedelta(days=6)
    return prev_monday, prev_sunday


def _week_utc_bounds(week_monday: date, week_sunday: date):
    start = IRELAND_TZ.localize(
        datetime(week_monday.year, week_monday.month, week_monday.day, 0, 0, 0)
    ).astimezone(pytz.utc)
    end = IRELAND_TZ.localize(
        datetime(week_sunday.year, week_sunday.month, week_sunday.day, 23, 59, 59)
    ).astimezone(pytz.utc)
    return start.isoformat(), end.isoformat()


def _get_team_member_name(member_id: str) -> str:
    try:
        r = requests.get(f'{SQUARE_BASE}/team-members/{member_id}',
                         headers=_headers(), timeout=10)
        if r.status_code == 200:
            m = r.json().get('team_member', {})
            return (m.get('display_name') or
                    f"{m.get('given_name','').strip()} {m.get('family_name','').strip()}".strip() or
                    member_id)
    except Exception as e:
        log.warning(f"Could not fetch team member {member_id}: {e}")
    return member_id


def get_shifts_for_week(target_sunday_override=None) -> dict:
    """
    Pull closed shifts for the Ireland week and filter manually by Ireland date.

    When a week crosses a month boundary (e.g. Mon Jul 27 - Sat Aug 1),
    hours are split:
      weekday_hours      = Mon-Sat shifts in the PRIMARY month (e.g. Mon-Fri Jul)
      sunday_hours       = Sunday shifts
      cross_month_hours  = weekday shifts in the OTHER month (e.g. Sat Aug 1)

    Returns:
        {sheet_name: {staff_name: {weekday_hours, sunday_hours, cross_month_hours}}}
    """
    week_monday, week_sunday = _get_week_range(target_sunday_override)
    start_at, end_at = _week_utc_bounds(week_monday, week_sunday)
    log.info(f"Collecting shifts for Ireland week: {week_monday} to {week_sunday}")

    # Determine if the weekday portion crosses a month boundary
    week_saturday = week_monday + timedelta(days=5)
    cross_month   = week_saturday.month != week_monday.month

    location_ids     = [lid for lid in LOCATION_MAP.values() if lid]
    location_reverse = {v: k for k, v in LOCATION_MAP.items() if v}

    if not location_ids:
        raise ValueError("No Square location IDs configured.")

    results      = {sheet: {} for sheet in LOCATION_MAP}
    name_cache   = {}
    cursor       = None
    total_shifts = 0
    used_shifts  = 0

    while True:
        body = {
            'filter': {
                'location_ids': location_ids,
                'start': {'start_at': start_at, 'end_at': end_at},
                'status': 'CLOSED',
            },
            'limit': 200,
        }
        if cursor:
            body['cursor'] = cursor

        r = requests.post(f'{SQUARE_BASE}/labor/shifts/search',
                          headers=_headers(), json=body)
        r.raise_for_status()
        data = r.json()

        for shift in data.get('shifts', []):
            total_shifts += 1
            location_id = shift.get('location_id')
            sheet_name  = location_reverse.get(location_id)
            if not sheet_name:
                continue

            start_str = shift.get('start_at')
            end_str   = shift.get('end_at')
            if not start_str or not end_str:
                continue

            shift_start = datetime.fromisoformat(start_str).astimezone(IRELAND_TZ)
            shift_end   = datetime.fromisoformat(end_str).astimezone(IRELAND_TZ)
            shift_date  = shift_start.date()

            if shift_date < week_monday or shift_date > week_sunday:
                continue

            used_shifts += 1
            hours = (shift_end - shift_start).total_seconds() / 3600

            member_id = shift.get('team_member_id', '')
            if member_id not in name_cache:
                name_cache[member_id] = _get_team_member_name(member_id)
            name = name_cache[member_id]

            if name not in results[sheet_name]:
                results[sheet_name][name] = {
                    'weekday_hours':      0.0,
                    'sunday_hours':       0.0,
                    'cross_month_hours':  0.0,
                }

            if shift_start.weekday() == 6:   # Sunday
                results[sheet_name][name]['sunday_hours'] += hours
            elif cross_month and shift_date.month != week_monday.month:
                # Shift is in the cross-month portion (e.g. Saturday of next month)
                results[sheet_name][name]['cross_month_hours'] += hours
            else:
                results[sheet_name][name]['weekday_hours'] += hours

        cursor = data.get('cursor')
        if not cursor:
            break

    log.info(f"Square returned {total_shifts} shifts total, {used_shifts} in target week")
    for sheet, staff in results.items():
        log.info(f"[{sheet}] {len(staff)} staff with shifts this week")

    return results


def _fetch_income_for_period(location_id: str, period_start: date, period_end: date) -> float:
    """
    Fetch Totals Collected - Gift Vouchers Redeemed for a specific date range and location.
    period_start and period_end are inclusive dates (Ireland time).
    """
    start_at, end_at = _week_utc_bounds(period_start, period_end)
    total_collected    = 0.0
    gift_card_redeemed = 0.0
    cursor = None

    while True:
        params = {
            'location_id': location_id,
            'begin_time':  start_at,
            'end_time':    end_at,
            'limit':       500,
        }
        if cursor:
            params['cursor'] = cursor

        r = requests.get(f'{SQUARE_BASE}/payments', headers=_headers(), params=params)
        r.raise_for_status()
        data = r.json()

        for payment in data.get('payments', []):
            amount = payment.get('total_money', {}).get('amount', 0) / 100
            total_collected += amount
            if payment.get('source_type') == 'SQUARE_GIFT_CARD':
                gift_card_redeemed += amount

        cursor = data.get('cursor')
        if not cursor:
            break

    return round(total_collected - gift_card_redeemed, 2)


def get_income_for_week(target_sunday_override=None) -> dict:
    """
    Fetch Totals Collected - Gift Vouchers Redeemed for each location.

    When a week crosses a month boundary, TWO separate Square API calls are made
    to get the actual income for each portion — no approximation.

        e.g. Mon-Fri Jul 27-31 → call Square for Jul 27-31 → exact July income
             Sat Aug 1          → call Square for Aug 1 only → exact August income

    Returns:
        {sheet_name: {'income': float, 'cross_month_income': float}}
    """
    week_monday, week_sunday = _get_week_range(target_sunday_override)
    log.info(f"Collecting income for Ireland week: {week_monday} to {week_sunday}")

    week_saturday = week_monday + timedelta(days=5)
    cross_month   = week_saturday.month != week_monday.month

    if cross_month:
        import calendar
        last_of_month  = date(week_monday.year, week_monday.month,
                              calendar.monthrange(week_monday.year, week_monday.month)[1])
        primary_end    = last_of_month     # e.g. Fri Jul 31
        cross_start    = week_saturday     # e.g. Sat Aug 1
        cross_end      = week_saturday     # same day
    else:
        primary_end = week_saturday
        cross_start = cross_end = None

    results = {}

    for sheet_name, location_id in LOCATION_MAP.items():
        if not location_id:
            results[sheet_name] = {'income': 0.0, 'cross_month_income': 0.0}
            continue

        income_primary = _fetch_income_for_period(location_id, week_monday, primary_end)

        if cross_month and cross_start:
            income_cross = _fetch_income_for_period(location_id, cross_start, cross_end)
        else:
            income_cross = 0.0

        results[sheet_name] = {
            'income':            income_primary,
            'cross_month_income': income_cross,
        }
        log.info(
            f"[{sheet_name}] Income: €{income_primary:.2f} ({week_monday} to {primary_end})"
            + (f" + €{income_cross:.2f} ({cross_start})" if income_cross else "")
        )

    return results
