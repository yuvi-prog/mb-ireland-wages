import os
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import pytz
from flask import Flask, request, jsonify, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from square_service import get_shifts_for_week, get_income_for_week, get_previous_week_sunday
from excel_service import update_excel_wages, write_prev_month_end
from email_service import send_wages_email, send_wages_email_cross_month
from month_service import (
    refresh_for_month, monthly_filename, MONTH_NAMES
)

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(name)s — %(message)s')
log = logging.getLogger(__name__)

app          = Flask(__name__)
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

def run_wages(target_sunday_override=None, end_of_month_preliminary=False, cutoff_date=None):
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
            prev_year, prev_month = week_monday.year, week_monday.month
            curr_year, curr_month = week_saturday.year, week_saturday.month

            prev_file = _get_or_create_file(prev_year, prev_month)
            curr_file = _get_or_create_file(curr_year, curr_month)

            # 1. Write Mon→last_day_of_prev_month hours to prev month's file
            result_prev = write_prev_month_end(str(prev_file), shifts, income)

            # 2. Write curr_month hours (cross_month_hours) + Sunday to curr month's file.
            #    Swap cross_month_hours → weekday_hours so they land in the first data column.
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
                    'income':            v.get('cross_month_income', 0.0),
                    'cross_month_income': 0.0,
                }
                for loc, v in income.items()
            }
            result_curr = update_excel_wages(str(curr_file), curr_shifts, curr_income, target_sunday)

            send_wages_email_cross_month(
                str(prev_file), str(curr_file),
                result_prev, result_curr,
                target_sunday,
                preliminary=end_of_month_preliminary
            )
        else:
            # Normal single-month week
            file_path = get_current_monthly_file(target_sunday)
            if not file_path:
                return
            result = update_excel_wages(str(file_path), shifts, income, target_sunday)
            send_wages_email(str(file_path), result, target_sunday, preliminary=end_of_month_preliminary)

        log.info("=== Weekly wages run complete ===")

    except Exception as e:
        log.error(f"Wages run failed: {e}", exc_info=True)


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
