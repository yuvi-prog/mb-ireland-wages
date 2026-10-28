import os
import logging
import threading
from datetime import datetime, date, timedelta
from pathlib import Path

import pytz
from flask import Flask, request, jsonify, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from square_service import get_shifts_for_week, get_income_for_week, get_previous_week_sunday
from excel_service import update_excel_wages, write_prev_month_end
from email_service import send_wages_email, send_wages_email_cross_month, send_catchup_email
from admin import admin_bp
from month_service import (
    refresh_for_month, monthly_filename, MONTH_NAMES
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s — %(message)s')
log = logging.getLogger(__name__)

app          = Flask(__name__)
app.register_blueprint(admin_bp)
DATA_DIR     = Path(os.getenv('DATA_DIR', '/data'))
API_KEY      = os.getenv('API_KEY', 'changeme')
MELBOURNE_TZ = pytz.timezone('Australia/Melbourne')

DATA_DIR.mkdir(parents=True, exist_ok=True)
MASTER_TEMPLATE = DATA_DIR / 'master_template.xlsx'


def _auth(req) -> bool:
    return req.headers.get('X-API-Key') == API_KEY or req.args.get('key') == API_KEY


def _get_or_create_file(year: int, month: int) -> Path:
    """Get or auto-generate a monthly wages file."""
    path = DATA_DIR / monthly_filename(year, month)
    if not path.exists():
        if not MASTER_TEMPLATE.exists():
            raise FileNotFoundError("Master template not found — upload via /upload-template")
        log.info(f"Auto-generating {path.name} from master template")
        refresh_for_month(str(MASTER_TEMPLATE), str(path), year, month)
    return path


def get_current_monthly_file(target_sunday=None) -> Path | None:
    try:
        if target_sunday:
            return _get_or_create_file(target_sunday.year, target_sunday.month)
        now = datetime.now(MELBOURNE_TZ)
        return _get_or_create_file(now.year, now.month)
    except FileNotFoundError as e:
        log.error(str(e))
        return None


# ── Main wages job ─────────────────────────────────────────────────────────────

# One run at a time: the Monday job, manual triggers and catch-up all write
# to the same Excel files, so they must not overlap.
_run_lock = threading.RLock()

MAX_CATCHUP_WEEKS = 13


def run_wages(target_sunday_override=None, end_of_month_preliminary=False,
              cutoff_date=None, send_email=True):
    """
    Process one payroll week and write it into the monthly Excel file(s).

    Returns {'ok': bool, 'files': [paths written], 'error': str | None}.
    send_email=False is used by catch-up, which sends one combined email at the end.
    """
    with _run_lock:
        log.info("=== Weekly wages run started ===")

        try:
            target_sunday = target_sunday_override or get_previous_week_sunday()
            week_monday   = target_sunday - timedelta(days=6)
            week_saturday = week_monday + timedelta(days=5)
            is_cross      = week_saturday.month != week_monday.month

            if end_of_month_preliminary:
                log.info(f"End-of-month preliminary run for {week_monday} to {target_sunday}")
            log.info(f"Processing week {week_monday} to {target_sunday} "
                     f"({'cross-month' if is_cross else 'single-month'})")

            shifts = get_shifts_for_week(target_sunday_override=target_sunday, cutoff_date=cutoff_date)
            income = get_income_for_week(target_sunday_override=target_sunday, cutoff_date=cutoff_date)

            if is_cross:
                # Cross-month week: write to TWO files
                prev_file = _get_or_create_file(week_monday.year, week_monday.month)
                curr_file = _get_or_create_file(week_saturday.year, week_saturday.month)

                # 1. Mon -> last day of previous month goes in the previous month's file
                result_prev = write_prev_month_end(str(prev_file), shifts, income)

                # 2. New-month days + Sunday go in the current month's file.
                #    Swap cross_month_hours into weekday_hours so they land in the first data column.
                curr_shifts = {
                    loc: {
                        name: {
                            'weekday_hours':     round(h.get('cross_month_hours', 0.0), 2),
                            'sunday_hours':      round(h['sunday_hours'], 2),
                            'cross_month_hours': 0.0,
                        }
                        for name, h in staff.items()
                    }
                    for loc, staff in shifts.items()
                }
                curr_income = {
                    loc: {
                        'income':             v.get('cross_month_income', 0.0),
                        'cross_month_income': 0.0,
                    }
                    for loc, v in income.items()
                }
                result_curr = update_excel_wages(str(curr_file), curr_shifts, curr_income, target_sunday)

                if send_email:
                    send_wages_email_cross_month(
                        str(prev_file), str(curr_file),
                        result_prev, result_curr,
                        target_sunday,
                        preliminary=end_of_month_preliminary
                    )
                files = [str(prev_file), str(curr_file)]
            else:
                # Normal single-month week
                # File is chosen by the Saturday's month: when the month ends on a Saturday the
                # Sunday is already next month, but that Sunday column lives in this month's file.
                file_path = get_current_monthly_file(week_saturday)
                if not file_path:
                    return {'ok': False, 'files': [], 'error': 'Master template not found'}
                result = update_excel_wages(str(file_path), shifts, income, target_sunday)
                if send_email:
                    send_wages_email(str(file_path), result, target_sunday,
                                     preliminary=end_of_month_preliminary)
                files = [str(file_path)]

            log.info("=== Weekly wages run complete ===")
            return {'ok': True, 'files': files, 'error': None}

        except Exception as e:
            log.error(f"Wages run failed: {e}", exc_info=True)
            return {'ok': False, 'files': [], 'error': str(e)}


# ── Catch-up ───────────────────────────────────────────────────────────────────

def _catchup_plan(today_ie: date, from_date: date = None) -> list:
    """
    Weeks to (re)build, as (monday, sunday, cutoff_date).

    cutoff_date is None for a finished week, or today's date for the current
    unfinished week so only shifts up to today are pulled.
    Defaults to starting at the first payroll week of the current month.
    """
    if from_date is None:
        from_date = today_ie.replace(day=1)
    monday = from_date - timedelta(days=from_date.weekday())

    plan = []
    while monday <= today_ie:
        sunday = monday + timedelta(days=6)
        plan.append((monday, sunday, None if sunday < today_ie else today_ie))
        monday += timedelta(days=7)
    return plan


def run_catchup(from_date=None, today_ie=None):
    """
    Rebuild every week from from_date (default: start of this month) up to today,
    then send ONE email with the updated file(s) and a week-by-week status list.
    A failed week is recorded and flagged in the email; the others still run.
    """
    today_ie = today_ie or datetime.now(pytz.timezone('Europe/Dublin')).date()
    plan     = _catchup_plan(today_ie, from_date)
    log.info(f"=== Catch-up started: {len(plan)} weeks, up to {today_ie} (Ireland) ===")

    statuses, files = [], []
    with _run_lock:
        for monday, sunday, cutoff in plan:
            res = run_wages(
                target_sunday_override=sunday,
                end_of_month_preliminary=bool(cutoff),
                cutoff_date=cutoff,
                send_email=False,
            )
            statuses.append({
                'monday':  monday,
                'sunday':  sunday,
                'partial': bool(cutoff),
                'ok':      res['ok'],
                'error':   res['error'],
            })
            for f in res['files']:
                if f not in files:
                    files.append(f)

    try:
        send_catchup_email(files, statuses, today_ie)
    except Exception as e:
        log.error(f"Catch-up email failed: {e}", exc_info=True)

    ok = sum(1 for s in statuses if s['ok'])
    log.info(f"=== Catch-up complete: {ok}/{len(statuses)} weeks OK ===")
    return statuses


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route('/health')
def health():
    f = get_current_monthly_file()
    return jsonify({
        'status':          'ok',
        'master_template': MASTER_TEMPLATE.exists(),
        'current_file':    f.name if f else None,
        'server_time':     datetime.now(MELBOURNE_TZ).isoformat(),
    })


@app.route('/catch-up', methods=['GET', 'POST'])
def catch_up():
    """
    Rebuild every week from the start of this month up to today and email the result.
    Optional ?from=YYYY-MM-DD to start earlier (snaps back to that week's Monday).
    Runs in the background so the request returns straight away.
    """
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    from_date = None
    from_param = request.args.get('from')
    if from_param:
        try:
            from_date = date.fromisoformat(from_param)
        except ValueError:
            return jsonify({'error': 'invalid date, use YYYY-MM-DD'}), 400

    today_ie = datetime.now(pytz.timezone('Europe/Dublin')).date()
    plan = _catchup_plan(today_ie, from_date)
    if not plan:
        return jsonify({'error': 'nothing to catch up, start date is in the future'}), 400
    if len(plan) > MAX_CATCHUP_WEEKS:
        return jsonify({'error': f'{len(plan)} weeks is too many, max is {MAX_CATCHUP_WEEKS}'}), 400

    threading.Thread(target=run_catchup, args=(from_date, today_ie), daemon=True).start()
    return jsonify({
        'status': 'started',
        'up_to':  str(today_ie),
        'weeks':  [f"{m} to {s}" + (" (so far)" if c else "") for m, s, c in plan],
        'note':   'The email arrives when every week has finished, usually 1-3 minutes.',
    })


@app.route('/trigger-preliminary', methods=['GET', 'POST'])
def trigger_preliminary():
    """
    Trigger a preliminary run up to today (Ireland time).
    Uses the current week's Sunday for correct column detection,
    but only fetches shifts from Monday up to today.
    """
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    ireland_tz  = pytz.timezone('Europe/Dublin')
    today_ie    = datetime.now(ireland_tz).date()
    dow         = today_ie.weekday()          # Mon=0
    week_monday = today_ie - timedelta(days=dow)
    week_sunday = week_monday + timedelta(days=6)

    log.info(f"Preliminary trigger: week {week_monday} to {week_sunday}, "
             f"fetching up to today {today_ie} (Ireland)")

    # Pass week_sunday as target so the correct column is found in the file,
    # but pass cutoff_date=today so Square only returns shifts up to now.
    run_wages(
        target_sunday_override=week_sunday,
        end_of_month_preliminary=True,
        cutoff_date=today_ie,
    )
    return jsonify({
        'status':   'triggered',
        'week':     f"{week_monday} to {week_sunday}",
        'up_to':    str(today_ie),
        'at':       datetime.now(MELBOURNE_TZ).isoformat(),
    })


@app.route('/trigger', methods=['GET', 'POST'])
def trigger():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    custom_date = request.args.get('date')
    if custom_date:
        try:
            override_monday = date.fromisoformat(custom_date)
            target_sunday   = override_monday - timedelta(days=1)
            log.info(f"Trigger override: Monday {custom_date} → Sunday {target_sunday}")
            run_wages(target_sunday_override=target_sunday)
        except ValueError:
            return jsonify({'error': 'invalid date, use YYYY-MM-DD'}), 400
    else:
        run_wages()

    return jsonify({'status': 'triggered', 'at': datetime.now(MELBOURNE_TZ).isoformat()})


@app.route('/upload-template', methods=['POST'])
def upload_template():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    if 'file' not in request.files:
        return jsonify({'error': 'no file (field: file)'}), 400
    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': 'must be .xlsx'}), 400

    f.save(MASTER_TEMPLATE)
    now  = datetime.now(MELBOURNE_TZ)
    path = DATA_DIR / monthly_filename(now.year, now.month)
    refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)
    log.info(f"Master template uploaded, generated {path.name}")
    return jsonify({'status': 'uploaded', 'generated': path.name})


@app.route('/regenerate', methods=['GET', 'POST'])
def regenerate():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    if not MASTER_TEMPLATE.exists():
        return jsonify({'error': 'no master template'}), 400

    month_param = request.args.get('month')
    if month_param:
        try:
            parts = month_param.split('-')
            year, month = int(parts[0]), int(parts[1])
        except (ValueError, IndexError):
            return jsonify({'error': 'invalid month format, use YYYY-MM'}), 400
    else:
        now   = datetime.now(MELBOURNE_TZ)
        year  = now.year
        month = now.month

    filename = monthly_filename(year, month)
    path     = DATA_DIR / filename
    refresh_for_month(str(MASTER_TEMPLATE), str(path), year, month)
    return jsonify({'status': 'regenerated', 'file': filename})


@app.route('/download')
def download():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    f = get_current_monthly_file()
    if not f:
        return jsonify({'error': 'no file'}), 404
    return send_file(f, as_attachment=True)


@app.route('/list-files')
def list_files():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    files = sorted(DATA_DIR.glob('MB_Ireland_*.xlsx'), reverse=True)
    return jsonify({'master_template': MASTER_TEMPLATE.exists(), 'monthly_files': [f.name for f in files]})


# ── Scheduler ──────────────────────────────────────────────────────────────────

scheduler = BackgroundScheduler(timezone=MELBOURNE_TZ)

# Weekly run — every Monday 8am Melbourne time
scheduler.add_job(
    run_wages,
    CronTrigger(day_of_week='mon', hour=8, minute=0, timezone=MELBOURNE_TZ),
    id='weekly_wages', replace_existing=True,
)

# End-of-month run — 1st of every month at 8am Melbourne time
# By then it's ~10-11pm the previous night in Ireland so almost all
# last-day shifts are clocked out. Email is flagged as preliminary.
scheduler.add_job(
    lambda: run_wages(end_of_month_preliminary=True),
    CronTrigger(day=1, hour=8, minute=0, timezone=MELBOURNE_TZ),
    id='end_of_month_wages', replace_existing=True,
)

scheduler.start()
log.info("Scheduler running — wages: Mon 08:00 + 1st of month 08:00 Melbourne time")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)
