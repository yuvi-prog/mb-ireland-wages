import os
import logging
from datetime import datetime
from pathlib import Path

import pytz
from flask import Flask, request, jsonify, send_file
from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.cron import CronTrigger

from square_service import get_shifts_for_week, get_previous_week_sunday
from excel_service import update_excel_wages
from email_service import send_wages_email
from month_service import (
    refresh_for_month, monthly_filename, current_expected_filename,
    MONTH_NAMES
)

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s — %(message)s',
)
log = logging.getLogger(__name__)

app        = Flask(__name__)
DATA_DIR   = Path(os.getenv('DATA_DIR', '/data'))
API_KEY    = os.getenv('API_KEY', 'changeme')
IRELAND_TZ = pytz.timezone('Europe/Dublin')

DATA_DIR.mkdir(parents=True, exist_ok=True)

MASTER_TEMPLATE = DATA_DIR / 'master_template.xlsx'


# ── helpers ───────────────────────────────────────────────────────────────────

def _auth(req) -> bool:
    return (
        req.headers.get('X-API-Key') == API_KEY
        or req.args.get('key') == API_KEY
    )


def get_current_monthly_file() -> Path | None:
    """
    Return the wages file for the current month.
    If it doesn't exist yet, auto-generate it from the master template.
    """
    now      = datetime.now(IRELAND_TZ)
    filename = monthly_filename(now.year, now.month)
    path     = DATA_DIR / filename

    if not path.exists():
        if not MASTER_TEMPLATE.exists():
            log.error("Master template not found — upload it via POST /upload-template")
            return None
        log.info(f"Generating {filename} from master template...")
        refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)

    return path


# ── main job ──────────────────────────────────────────────────────────────────

def run_wages():
    log.info("=== Weekly wages run started ===")

    file_path = get_current_monthly_file()
    if not file_path:
        return

    log.info(f"Using file: {file_path.name}")

    try:
        target_sunday = get_previous_week_sunday()
        log.info(f"Processing week ending Sunday {target_sunday}")

        shifts = get_shifts_for_week()
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
        'server_time':     datetime.now(IRELAND_TZ).isoformat(),
    })


@app.route('/trigger', methods=['GET', 'POST'])
def trigger():
    """Manually trigger a wages run."""
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401
    run_wages()
    return jsonify({'status': 'triggered', 'at': datetime.now(IRELAND_TZ).isoformat()})


@app.route('/upload-template', methods=['POST'])
def upload_template():
    """
    Upload the master template (do this ONCE, or whenever staff/rates change).
    This is the blank wages file with all staff names, rates, and formulas.

    Usage:
        curl -X POST https://your-app.railway.app/upload-template?key=YOUR_KEY \\
             -F "file=@MB_Ireland_Jun_2026.xlsx"
    """
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    if 'file' not in request.files:
        return jsonify({'error': 'no file in request (field name: file)'}), 400

    f = request.files['file']
    if not f.filename.endswith('.xlsx'):
        return jsonify({'error': 'must be .xlsx'}), 400

    f.save(MASTER_TEMPLATE)
    log.info("Master template uploaded")

    # Pre-generate this month's file right away
    now  = datetime.now(IRELAND_TZ)
    path = DATA_DIR / monthly_filename(now.year, now.month)
    refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)
    log.info(f"Auto-generated {path.name} from new master template")

    return jsonify({
        'status':          'uploaded',
        'master_template': 'master_template.xlsx',
        'generated':       path.name,
    })


@app.route('/regenerate', methods=['POST'])
def regenerate():
    """
    Force-regenerate the current month's file from the master template.
    Useful if you've updated staff or rates in the master and want to start fresh.
    WARNING: this wipes any hours already entered for the current month.
    """
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    if not MASTER_TEMPLATE.exists():
        return jsonify({'error': 'no master template — upload one first'}), 400

    now      = datetime.now(IRELAND_TZ)
    filename = monthly_filename(now.year, now.month)
    path     = DATA_DIR / filename

    refresh_for_month(str(MASTER_TEMPLATE), str(path), now.year, now.month)
    return jsonify({'status': 'regenerated', 'file': filename})


@app.route('/download')
def download():
    """Download the current month's wages file."""
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    f = get_current_monthly_file()
    if not f:
        return jsonify({'error': 'no file — upload master template first'}), 404

    return send_file(f, as_attachment=True)


@app.route('/list-files')
def list_files():
    """List all generated monthly files."""
    if not _auth(request):
        return jsonify({'error': 'unauthorized'}), 401

    files = sorted(DATA_DIR.glob('MB_Ireland_*.xlsx'), reverse=True)
    return jsonify({
        'master_template': MASTER_TEMPLATE.exists(),
        'monthly_files':   [f.name for f in files],
    })


# ── scheduler ─────────────────────────────────────────────────────────────────

scheduler = BackgroundScheduler(timezone=IRELAND_TZ)
scheduler.add_job(
    run_wages,
    CronTrigger(day_of_week='mon', hour=8, minute=0, timezone=IRELAND_TZ),
    id='weekly_wages',
    replace_existing=True,
)
scheduler.start()
log.info("Scheduler running — wages job fires Monday 08:00 Ireland time")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=False)
