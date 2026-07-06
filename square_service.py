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
    """Return (monday, sunday) as date objects for the target week."""
    prev_sunday = target_sunday or get_previous_week_sunday()
    prev_monday = prev_sunday - timedelta(days=6)
    return prev_monday, prev_sunday


def _get_team_member_name(member_id: str) -> str:
    try:
        r = requests.get(
            f'{SQUARE_BASE}/team-members/{member_id}',
            headers=_headers(),
            timeout=10,
        )
        if r.status_code == 200:
            m = r.json().get('team_member', {})
            return (
                m.get('display_name')
                or f"{m.get('given_name','').strip()} {m.get('family_name','').strip()}".strip()
                or member_id
            )
    except Exception as e:
        log.warning(f"Could not fetch team member {member_id}: {e}")
    return member_id


def get_shifts_for_week(target_sunday_override=None) -> dict:
    """
    Pull shifts from Square and filter manually by shift start date in Ireland time.
    Square's date filter is unreliable so we fetch a broader window and filter ourselves.
    """
    week_monday, week_sunday = _get_week_range(target_sunday_override)
    log.info(f"Collecting shifts for Ireland week: {week_monday} to {week_sunday}")

    # Fetch a 10-day window to be safe (covers timezone edge cases)
    fetch_start = IRELAND_TZ.localize(
        datetime(week_monday.year, week_monday.month, week_monday.day, 0, 0, 0)
    ).astimezone(pytz.utc)
    fetch_end = IRELAND_TZ.localize(
        datetime(week_sunday.year, week_sunday.month, week_sunday.day, 23, 59, 59)
    ).astimezone(pytz.utc)

    location_ids     = [lid for lid in LOCATION_MAP.values() if lid]
    location_reverse = {v: k for k, v in LOCATION_MAP.items() if v}

    if not location_ids:
        raise ValueError("No Square location IDs configured.")

    results    = {sheet: {} for sheet in LOCATION_MAP}
    name_cache = {}
    cursor     = None
    total_shifts = 0
    filtered_shifts = 0

    while True:
        body = {
            'filter': {
                'location_ids': location_ids,
                'start': {
                    'start_at': fetch_start.isoformat(),
                    'end_at':   fetch_end.isoformat(),
                },
                'status': 'CLOSED',
            },
            'limit': 200,
        }
        if cursor:
            body['cursor'] = cursor

        r = requests.post(f'{SQUARE_BASE}/labor/shifts/search', headers=_headers(), json=body)
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

            # Convert to Ireland time and manually check the shift is in our week
            shift_start_ireland = datetime.fromisoformat(start_str).astimezone(IRELAND_TZ)
            shift_end_ireland   = datetime.fromisoformat(end_str).astimezone(IRELAND_TZ)
            shift_date          = shift_start_ireland.date()

            if shift_date < week_monday or shift_date > week_sunday:
                continue  # outside our target week — skip

            filtered_shifts += 1
            hours = (shift_end_ireland - shift_start_ireland).total_seconds() / 3600

            member_id = shift.get('team_member_id', '')
            if member_id not in name_cache:
                name_cache[member_id] = _get_team_member_name(member_id)
            name = name_cache[member_id]

            if name not in results[sheet_name]:
                results[sheet_name][name] = {'weekday_hours': 0.0, 'sunday_hours': 0.0}

            if shift_start_ireland.weekday() == 6:  # Sunday
                results[sheet_name][name]['sunday_hours']  += hours
            else:
                results[sheet_name][name]['weekday_hours'] += hours

        cursor = data.get('cursor')
        if not cursor:
            break

    log.info(f"Square returned {total_shifts} shifts total, {filtered_shifts} in target week")
    for sheet, staff in results.items():
        log.info(f"[{sheet}] {len(staff)} staff with shifts this week")

    return results
