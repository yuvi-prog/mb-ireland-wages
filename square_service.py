import os
import logging
from datetime import datetime, timedelta, date

import pytz
import requests

log = logging.getLogger(__name__)

MELBOURNE_TZ = pytz.timezone('Australia/Melbourne')
SQUARE_BASE = 'https://connect.squareup.com/v2'

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


def _get_week_range(target_sunday=None) -> tuple:
    prev_sunday = target_sunday if target_sunday else get_previous_week_sunday()
    prev_monday = prev_sunday - timedelta(days=6)
    monday_dt = MELBOURNE_TZ.localize(datetime(prev_monday.year, prev_monday.month, prev_monday.day, 0, 0, 0))
    sunday_dt  = MELBOURNE_TZ.localize(datetime(prev_sunday.year,  prev_sunday.month,  prev_sunday.day,  23, 59, 59))
    return monday_dt.isoformat(), sunday_dt.isoformat()


def _get_team_member_name(member_id: str) -> str:
    """
    Fetch a single team member's name by ID.
    Falls back to the member ID string if the API call fails.
    """
    try:
        r = requests.get(
            f'{SQUARE_BASE}/team-members/{member_id}',
            headers=_headers(),
            timeout=10,
        )
        if r.status_code == 200:
            m = r.json().get('team_member', {})
            return m.get('display_name') or f"{m.get('given_name','').strip()} {m.get('family_name','').strip()}".strip() or member_id
    except Exception as e:
        log.warning(f"Could not fetch team member {member_id}: {e}")
    return member_id


def get_shifts_for_week(target_sunday_override=None) -> dict:
    """
    Pull all closed shifts for the previous week across all Irish locations.

    Returns:
        {
            'Blanchardstown': {
                'Anron': {'weekday_hours': 15.5, 'sunday_hours': 8.0},
                ...
            },
            ...
        }
    """
    start_at, end_at = _get_week_range(target_sunday_override)
    log.info(f"Fetching shifts from {start_at} to {end_at}")

    location_ids    = [lid for lid in LOCATION_MAP.values() if lid]
    location_reverse = {v: k for k, v in LOCATION_MAP.items() if v}

    if not location_ids:
        raise ValueError("No Square location IDs configured.")

    results    = {sheet: {} for sheet in LOCATION_MAP}
    name_cache = {}   # member_id -> display name
    cursor     = None

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

        r = requests.post(f'{SQUARE_BASE}/labor/shifts/search', headers=_headers(), json=body)
        r.raise_for_status()
        data = r.json()

        for shift in data.get('shifts', []):
            location_id = shift.get('location_id')
            sheet_name  = location_reverse.get(location_id)
            if not sheet_name:
                continue

            member_id = shift.get('team_member_id', '')

            # Cache name lookups so we only call the API once per person
            if member_id not in name_cache:
                name_cache[member_id] = _get_team_member_name(member_id)
            name = name_cache[member_id]

            start_str = shift.get('start_at')
            end_str   = shift.get('end_at')
            if not start_str or not end_str:
                continue

            shift_start = datetime.fromisoformat(start_str).astimezone(MELBOURNE_TZ)
            shift_end   = datetime.fromisoformat(end_str).astimezone(MELBOURNE_TZ)
            hours       = (shift_end - shift_start).total_seconds() / 3600

            if name not in results[sheet_name]:
                results[sheet_name][name] = {'weekday_hours': 0.0, 'sunday_hours': 0.0}

            if shift_start.weekday() == 6:
                results[sheet_name][name]['sunday_hours']  += hours
            else:
                results[sheet_name][name]['weekday_hours'] += hours

        cursor = data.get('cursor')
        if not cursor:
            break

    for sheet, staff in results.items():
        log.info(f"[{sheet}] {len(staff)} staff with shifts this week")

    return results
