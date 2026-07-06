import os
import logging
from datetime import datetime, date, timedelta
from pathlib import Path

import pytz
from flask import Flask, request, jsonify, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from square_service import get_shifts_for_week, get_previous_week_sunday
from excel_service import update_excel_wages
from email_service import send_wages_email
from month_service import (
    refresh_for_month, monthly_filename, MONTH_NAMES
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s — %(message)s',
)
log = logging.getLogger(__name__)

app        = Flask(__name__)
DATA_DIR   = Path(os.getenv('DATA_DIR', '/data'))
API_KEY    = os.getenv('API_KEY', 'changeme')
MELBOURNE_TZ = pytz.timezone('Australia/Melbourne')

DATA_DIR.mkdir(parents=True, exist_ok=True)

MASTER_TEMPLATE = DATA_DIR / 'master_template.xlsx'


def _auth(req) -> bool:
    return (
        req.headers.get('X-API-Key') == API_KEY
        or req.args.get('key') == API_KEY
    )


def get_current_monthly_file(target_sunday=None) -> Path | None:
    """Return the wages file for the month containing target_sunday (defaults to current month)."""
    if target_sunday:
        now_year  = target_sunday.year
        now_month = target_sunday.month
    else:
        now       = datetime.now(MELBOURNE_TZ)
        now_year  = now.year
        now_month = now.month

    filename = monthly_filename(now_year, now_month)
    path     = DATA_DIR / filename

    if not path.exists():
        if not MASTER_TEMPLATE.exists():
            log.error("Master template not found — upload via POST /upload-template")
            return None
        log.info(f"Generating {filename} from master template...")
        refresh_for_month(str(MASTER_TEMPLATE), str(path), now_year, now_month)

    return path


def run_wages(target_sunday_override=None):
    log.info("=== Weekly wages run started ===")

    target_sunday = target_sunday_override or get_previous_week_sunday()
    file_path     = get_current_monthly_file(target_sunday)

    if not file_path:
        return

    log.info(f"Using file: {file_path.name}")
    log.info(f"Processing week ending Sunday {target_sunday}")

    try:
        shifts = get_shifts_for_week(target_sunday_override=target_sunday)
        result = update_excel_wages(str(file_path), shifts, target_sunday)
        send_wages_email(str(file_path), result, target_sunday)
        log.info("=== Weekly wages run complete ===")
    except Exception as e:
        log.error(f"Wages run failed: {e}", exc_info=True)


# ── routes ────────────────────────────────────────────────────────────────────

@app.route('/health')
def health():
    f = get_current_monthly_file() if MASTER_TEMPLATE.exists() else None
    return jsonify({
        'status':          'ok',
        'master_template': MASTER_TEMPLATE.exists(),
        'current_file':    f.name if f else None,
        'server_time':     datetime.now(MELBOURNE_TZ).isoformat(),
    })


@app.route('/trigger', methods=['GET', 'POST'])
def trigger():
    """
    Trigger a wages run.
    Pass ?date=2026-07-06 to simulate running on a specific Monday.
    """
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    custom_date = request.args.get('date')
    if custom_date:
        try:
            override_monday = date.fromisoformat(custom_date)
            target_sunday   = override_monday - timedelta(days=1)
            log.info(f"Trigger override: simulating Monday {custom_date}, target Sunday = {target_sunday}")
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
    log.info("Master template uploaded")

    now  = datetime.now(MELBOURNE_TZ)
    path = DATA_DIR / monthly_filename(now.year, now.month)
    refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)
    log.info(f"Auto-generated {path.name}")

    return jsonify({'status': 'uploaded', 'master_template': 'master_template.xlsx', 'generated': path.name})


@app.route('/regenerate', methods=['GET', 'POST'])
def regenerate():
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    if not MASTER_TEMPLATE.exists():
        return jsonify({'error': 'no master template'}), 400

    now      = datetime.now(MELBOURNE_TZ)
    filename = monthly_filename(now.year, now.month)
    path     = DATA_DIR / filename
    refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)
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


# ── scheduler ─────────────────────────────────────────────────────────────────

scheduler = BackgroundScheduler(timezone=MELBOURNE_TZ)
scheduler.add_job(
    run_wages,
    CronTrigger(day_of_week='mon', hour=8, minute=0, timezone=MELBOURNE_TZ),
    id='weekly_wages',
    replace_existing=True,
)
scheduler.start()
log.info("Scheduler running — wages job fires Monday 08:00 Ireland time")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)
