import os
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import pytz
from flask import Flask, request, jsonify, send_file, Response
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

# Tracks the outcome of the most recent wages run, surfaced on the dashboard.
last_run_info = {
    'timestamp': None,
    'status':    'never run',
}


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
        last_run_info['timestamp'] = datetime.now(MELBOURNE_TZ).isoformat()
        last_run_info['status']    = 'success'

    except Exception as e:
        log.error(f"Wages run failed: {e}", exc_info=True)
        last_run_info['timestamp'] = datetime.now(MELBOURNE_TZ).isoformat()
        last_run_info['status']    = f'failed: {e}'


# ── Routes ─────────────────────────────────────────────────────────────────────

def _next_scheduled_runs():
    """Compute the next weekly and next end-of-month scheduled run times."""
    now = datetime.now(MELBOURNE_TZ)

    # Next Monday 8am
    days_ahead = (0 - now.weekday()) % 7  # Monday = 0
    next_monday = (now + timedelta(days=days_ahead)).replace(hour=8, minute=0, second=0, microsecond=0)
    if next_monday <= now:
        next_monday += timedelta(days=7)

    # Next 1st-of-month 8am
    if now.day == 1 and now.replace(hour=8, minute=0, second=0, microsecond=0) > now:
        next_first = now.replace(hour=8, minute=0, second=0, microsecond=0)
    else:
        if now.month == 12:
            year, month = now.year + 1, 1
        else:
            year, month = now.year, now.month + 1
        next_first = MELBOURNE_TZ.localize(datetime(year, month, 1, 8, 0, 0))

    return next_monday, next_first


@app.route('/')
def dashboard():
    f = get_current_monthly_file()
    monthly_files = sorted(DATA_DIR.glob('MB_Ireland_*.xlsx'), reverse=True)
    next_monday, next_first = _next_scheduled_runs()

    last_run_ts     = last_run_info['timestamp'] or 'Never'
    last_run_status = last_run_info['status']

    master_status = (
        '<span style="color:#2e7d32;">✔ Uploaded</span>'
        if MASTER_TEMPLATE.exists()
        else '<span style="color:#c62828;">✘ Needs upload</span>'
    )

    links = [
        ('/health', 'GET', 'Check system health and current file status'),
        ('/trigger', 'GET/POST', 'Manually trigger the weekly wages run'),
        ('/trigger-preliminary', 'GET/POST', 'Trigger a preliminary run up to today'),
        ('/upload-template', 'POST', 'Upload the master wages template (.xlsx)'),
        ('/regenerate', 'GET/POST', 'Regenerate a monthly file from the master template'),
        ('/download', 'GET', 'Download the current month\'s wages file'),
        ('/list-files', 'GET', 'List the master template and all monthly files'),
    ]
    links_html = ''.join(
        f'<li><code>{path}</code> <span class="method">{method}</span> — {desc}</li>'
        for path, method, desc in links
    )

    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Ireland Wages Dashboard</title>
    <style>
        body {{
            font-family: -apple-system, Segoe UI, Roboto, Arial, sans-serif;
            max-width: 800px;
            margin: 40px auto;
            padding: 0 20px;
            color: #222;
            background: #fafafa;
        }}
        h1 {{ margin-bottom: 4px; }}
        .subtitle {{ color: #555; margin-top: 0; }}
        .card {{
            background: #fff;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            padding: 16px 20px;
            margin-bottom: 20px;
        }}
        .status-online {{ color: #2e7d32; font-weight: bold; }}
        table {{ width: 100%; border-collapse: collapse; }}
        td {{ padding: 6px 0; border-bottom: 1px solid #eee; }}
        td:first-child {{ color: #666; width: 45%; }}
        ul {{ list-style: none; padding: 0; margin: 0; }}
        li {{ padding: 6px 0; border-bottom: 1px solid #eee; }}
        li:last-child {{ border-bottom: none; }}
        code {{ background: #f0f0f0; padding: 2px 6px; border-radius: 4px; }}
        .method {{ font-size: 0.75em; color: #888; }}
        footer {{ color: #999; font-size: 0.85em; text-align: center; margin-top: 30px; }}
    </style>
</head>
<body>
    <h1>Ireland Wages System</h1>
    <p class="subtitle">
        Automates weekly wage calculations for Ireland-based staff by pulling shift and income
        data from Square, writing it into monthly wages spreadsheets, and emailing the results.
        Runs automatically every Monday and at the start of each month (Melbourne time).
    </p>

    <div class="card">
        <h2>System Status</h2>
        <table>
            <tr><td>Status</td><td class="status-online">● Online</td></tr>
            <tr><td>Server time (Melbourne)</td><td>{datetime.now(MELBOURNE_TZ).strftime('%Y-%m-%d %H:%M:%S %Z')}</td></tr>
            <tr><td>Last run</td><td>{last_run_ts}</td></tr>
            <tr><td>Last run status</td><td>{last_run_status}</td></tr>
            <tr><td>Next weekly run (Mon 8am)</td><td>{next_monday.strftime('%Y-%m-%d %H:%M %Z')}</td></tr>
            <tr><td>Next month-end run (1st 8am)</td><td>{next_first.strftime('%Y-%m-%d %H:%M %Z')}</td></tr>
        </table>
    </div>

    <div class="card">
        <h2>Wages Files</h2>
        <table>
            <tr><td>Master template</td><td>{master_status}</td></tr>
            <tr><td>Current month file</td><td>{f.name if f else 'Not available'}</td></tr>
            <tr><td>Monthly files stored</td><td>{len(monthly_files)}</td></tr>
        </table>
    </div>

    <div class="card">
        <h2>Quick Links</h2>
        <ul>
            {links_html}
        </ul>
    </div>

    <footer>Ireland Wages System</footer>
</body>
</html>
"""
    return Response(html, mimetype='text/html')


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
