"""
admin.py — Simple admin panel for MB Ireland Wages
Serves a password-protected HTML page for managing the wages system.
Add this to app.py with: from admin import admin_bp; app.register_blueprint(admin_bp)
"""

from flask import Blueprint, request, jsonify, make_response, redirect, url_for
from openpyxl import load_workbook
import os
from pathlib import Path
from functools import wraps

admin_bp = Blueprint('admin', __name__)

DATA_DIR        = Path(os.getenv('DATA_DIR', '/data'))
MASTER_TEMPLATE = DATA_DIR / 'master_template.xlsx'
ADMIN_PASSWORD  = os.getenv('API_KEY', 'changeme')

LOCATION_SHEETS = ['Blanchardstown', 'Cork', 'Liffey Valley', 'Nutgrove', 'Whitewater']
COL_NAME     = 3
COL_RATE_MON = 4
COL_RATE_SUN = 5


def check_auth(req):
    return req.cookies.get('admin_token') == ADMIN_PASSWORD


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not check_auth(request):
            return redirect('/admin/login')
        return f(*args, **kwargs)
    return decorated


def get_staff_from_template():
    """Read staff names and rates from master template."""
    if not MASTER_TEMPLATE.exists():
        return {}
    wb = load_workbook(MASTER_TEMPLATE)
    result = {}
    for sheet in LOCATION_SHEETS:
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        staff = []
        for row in ws.iter_rows(min_row=7, max_row=50, min_col=COL_NAME, max_col=COL_RATE_SUN):
            name = row[0].value
            if not name or not isinstance(name, str) or not name.strip():
                continue
            staff.append({
                'name':     name.strip(),
                'rate_mon': row[1].value or '',
                'rate_sun': row[2].value or '',
            })
        result[sheet] = staff
    return result


def save_staff_to_template(location, staff_list):
    """Update staff names and rates in master template for a location."""
    if not MASTER_TEMPLATE.exists():
        return False, 'Master template not found'
    wb = load_workbook(MASTER_TEMPLATE)
    if location not in wb.sheetnames:
        return False, f'Sheet {location} not found'
    ws = wb[location]

    # Clear existing staff rows
    for row in ws.iter_rows(min_row=7, max_row=50, min_col=COL_NAME, max_col=COL_RATE_SUN):
        for cell in row:
            if not isinstance(cell.value, str) or not cell.value.startswith('='):
                cell.value = None

    # Write updated staff list
    for i, staff in enumerate(staff_list):
        r = 7 + i
        ws.cell(r, COL_NAME).value     = staff['name']
        ws.cell(r, COL_RATE_MON).value = float(staff['rate_mon']) if staff['rate_mon'] else None
        ws.cell(r, COL_RATE_SUN).value = float(staff['rate_sun']) if staff['rate_sun'] else None

    wb.save(MASTER_TEMPLATE)
    return True, 'Saved'


@admin_bp.route('/admin/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        pwd = request.form.get('password', '')
        if pwd == ADMIN_PASSWORD:
            resp = make_response(redirect('/admin'))
            resp.set_cookie('admin_token', pwd, max_age=86400 * 7, httponly=True)
            return resp
        return login_page(error='Incorrect password')
    return login_page()


def login_page(error=''):
    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MB Ireland — Admin</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Inter', sans-serif; background: #0f1117; color: #e2e8f0; min-height: 100vh;
         display: flex; align-items: center; justify-content: center; }}
  .card {{ background: #1a1f2e; border: 1px solid #2d3748; border-radius: 12px;
           padding: 40px; width: 100%; max-width: 380px; }}
  .logo {{ font-size: 13px; font-weight: 600; letter-spacing: 0.12em; text-transform: uppercase;
           color: #64748b; margin-bottom: 28px; }}
  h1 {{ font-size: 22px; font-weight: 600; color: #f1f5f9; margin-bottom: 8px; }}
  p {{ font-size: 14px; color: #64748b; margin-bottom: 28px; }}
  label {{ display: block; font-size: 12px; font-weight: 500; color: #94a3b8;
           text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 8px; }}
  input {{ width: 100%; padding: 11px 14px; background: #0f1117; border: 1px solid #2d3748;
           border-radius: 8px; color: #f1f5f9; font-size: 15px; outline: none; }}
  input:focus {{ border-color: #6366f1; }}
  button {{ width: 100%; margin-top: 20px; padding: 12px; background: #6366f1; color: #fff;
            border: none; border-radius: 8px; font-size: 15px; font-weight: 500; cursor: pointer; }}
  button:hover {{ background: #4f46e5; }}
  .error {{ background: #2d1f1f; border: 1px solid #7f1d1d; border-radius: 8px;
            padding: 10px 14px; font-size: 13px; color: #fca5a5; margin-bottom: 20px; }}
</style>
</head>
<body>
<div class="card">
  <div class="logo">Memory Block Ireland</div>
  <h1>Wages Admin</h1>
  <p>Sign in to manage staff and settings.</p>
  {"<div class='error'>" + error + "</div>" if error else ""}
  <form method="POST">
    <label>Password</label>
    <input type="password" name="password" placeholder="Enter admin password" autofocus>
    <button type="submit">Sign in</button>
  </form>
</div>
</body>
</html>'''


@admin_bp.route('/admin/logout')
def logout():
    resp = make_response(redirect('/admin/login'))
    resp.delete_cookie('admin_token')
    return resp


@admin_bp.route('/admin')
@login_required
def admin_home():
    staff_data = get_staff_from_template()
    template_exists = MASTER_TEMPLATE.exists()

    # Build monthly files list
    from pathlib import Path
    files = sorted(Path(DATA_DIR).glob('MB_Ireland_*.xlsx'), reverse=True)
    files_html = ''.join(
        f'<a href="/download?key={ADMIN_PASSWORD}" class="file-link">{f.name}</a>'
        for f in files[:3]
    )

    # Build staff tables per location
    location_tabs = ''
    location_panels = ''
    for i, loc in enumerate(LOCATION_SHEETS):
        staff = staff_data.get(loc, [])
        active = 'active' if i == 0 else ''
        location_tabs += f'<button class="tab {active}" onclick="showTab(\'{loc}\')" id="tab-{loc}">{loc}</button>'

        rows = ''
        for j, s in enumerate(staff):
            rows += f'''
            <tr>
              <td><input class="cell-input" name="name" value="{s["name"]}" placeholder="Name"></td>
              <td><input class="cell-input rate" name="rate_mon" value="{s["rate_mon"]}" placeholder="0.00"></td>
              <td><input class="cell-input rate" name="rate_sun" value="{s["rate_sun"]}" placeholder="0.00"></td>
              <td><button class="btn-remove" onclick="removeRow(this)">Remove</button></td>
            </tr>'''

        location_panels += f'''
        <div class="tab-panel {active}" id="panel-{loc}">
          <div class="panel-header">
            <h3>{loc}</h3>
            <button class="btn-add" onclick="addRow('{loc}')">+ Add staff member</button>
          </div>
          <table class="staff-table" id="table-{loc}">
            <thead>
              <tr>
                <th>Name</th>
                <th>Mon–Sat Rate (€/h)</th>
                <th>Sunday Rate (€/h)</th>
                <th></th>
              </tr>
            </thead>
            <tbody>{rows}</tbody>
          </table>
          <div class="save-row">
            <button class="btn-save" onclick="saveLocation('{loc}')">Save {loc}</button>
            <span class="save-status" id="status-{loc}"></span>
          </div>
        </div>'''

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>MB Ireland — Wages Admin</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Inter', sans-serif; background: #0f1117; color: #e2e8f0; min-height: 100vh; }}

  .topbar {{ background: #1a1f2e; border-bottom: 1px solid #2d3748; padding: 0 32px;
             display: flex; align-items: center; justify-content: space-between; height: 56px; }}
  .topbar-left {{ display: flex; align-items: center; gap: 16px; }}
  .logo {{ font-size: 12px; font-weight: 600; letter-spacing: 0.12em; text-transform: uppercase; color: #6366f1; }}
  .topbar h1 {{ font-size: 16px; font-weight: 600; color: #f1f5f9; }}
  .topbar-right {{ display: flex; gap: 10px; align-items: center; }}
  a.logout {{ font-size: 13px; color: #64748b; text-decoration: none; }}
  a.logout:hover {{ color: #94a3b8; }}

  .main {{ max-width: 1000px; margin: 0 auto; padding: 32px 24px; }}

  .cards {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 16px; margin-bottom: 32px; }}
  .card {{ background: #1a1f2e; border: 1px solid #2d3748; border-radius: 10px; padding: 20px; }}
  .card-label {{ font-size: 11px; font-weight: 600; text-transform: uppercase;
                 letter-spacing: 0.1em; color: #64748b; margin-bottom: 12px; }}
  .card-action {{ font-size: 14px; color: #94a3b8; margin-bottom: 14px; line-height: 1.5; }}
  .btn {{ display: inline-block; padding: 9px 16px; border-radius: 7px; font-size: 13px;
          font-weight: 500; cursor: pointer; border: none; text-decoration: none; }}
  .btn-primary {{ background: #6366f1; color: #fff; }}
  .btn-primary:hover {{ background: #4f46e5; }}
  .btn-secondary {{ background: #1e293b; border: 1px solid #334155; color: #94a3b8; }}
  .btn-secondary:hover {{ background: #263244; color: #e2e8f0; }}
  .btn-danger {{ background: #7f1d1d; color: #fca5a5; border: 1px solid #991b1b; }}
  .btn-danger:hover {{ background: #991b1b; }}

  .section {{ background: #1a1f2e; border: 1px solid #2d3748; border-radius: 10px;
              padding: 24px; margin-bottom: 24px; }}
  .section-title {{ font-size: 15px; font-weight: 600; color: #f1f5f9; margin-bottom: 20px; }}

  .tabs {{ display: flex; gap: 4px; margin-bottom: 20px; flex-wrap: wrap; }}
  .tab {{ padding: 7px 14px; border-radius: 6px; font-size: 13px; font-weight: 500;
          cursor: pointer; border: 1px solid #2d3748; background: transparent; color: #64748b; }}
  .tab:hover {{ color: #94a3b8; background: #1e293b; }}
  .tab.active {{ background: #6366f1; color: #fff; border-color: #6366f1; }}

  .tab-panel {{ display: none; }}
  .tab-panel.active {{ display: block; }}
  .panel-header {{ display: flex; align-items: center; justify-content: space-between; margin-bottom: 16px; }}
  .panel-header h3 {{ font-size: 15px; font-weight: 600; color: #e2e8f0; }}

  .staff-table {{ width: 100%; border-collapse: collapse; }}
  .staff-table th {{ text-align: left; font-size: 11px; font-weight: 600; text-transform: uppercase;
                     letter-spacing: 0.08em; color: #64748b; padding: 0 10px 10px 0; }}
  .staff-table td {{ padding: 4px 8px 4px 0; vertical-align: middle; }}
  .cell-input {{ background: #0f1117; border: 1px solid #2d3748; border-radius: 6px;
                 color: #e2e8f0; padding: 7px 10px; font-size: 14px; width: 100%; outline: none; }}
  .cell-input:focus {{ border-color: #6366f1; }}
  .cell-input.rate {{ max-width: 120px; }}
  .btn-remove {{ background: none; border: none; color: #64748b; cursor: pointer;
                 font-size: 18px; padding: 4px 8px; border-radius: 4px; }}
  .btn-remove:hover {{ color: #ef4444; background: #1f1515; }}
  .btn-add {{ background: none; border: 1px dashed #334155; color: #64748b; border-radius: 7px;
              padding: 7px 14px; font-size: 13px; cursor: pointer; }}
  .btn-add:hover {{ border-color: #6366f1; color: #6366f1; }}
  .save-row {{ display: flex; align-items: center; gap: 16px; margin-top: 20px;
               padding-top: 16px; border-top: 1px solid #1e293b; }}
  .btn-save {{ background: #6366f1; color: #fff; border: none; border-radius: 7px;
               padding: 9px 20px; font-size: 14px; font-weight: 500; cursor: pointer; }}
  .btn-save:hover {{ background: #4f46e5; }}
  .save-status {{ font-size: 13px; color: #22c55e; }}

  .upload-form {{ display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }}
  .file-input {{ background: #0f1117; border: 1px solid #2d3748; border-radius: 7px;
                 color: #94a3b8; padding: 8px 12px; font-size: 13px; cursor: pointer; }}
  .file-link {{ display: block; font-size: 13px; color: #6366f1; text-decoration: none;
                margin-top: 8px; }}
  .file-link:hover {{ text-decoration: underline; }}

  .toast {{ position: fixed; bottom: 24px; right: 24px; background: #22c55e; color: #fff;
            padding: 12px 20px; border-radius: 8px; font-size: 14px; font-weight: 500;
            opacity: 0; transition: opacity 0.3s; pointer-events: none; z-index: 999; }}
  .toast.show {{ opacity: 1; }}
  .toast.error {{ background: #ef4444; }}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-left">
    <span class="logo">MB Ireland</span>
    <h1>Wages Admin</h1>
  </div>
  <div class="topbar-right">
    <a href="/admin/logout" class="logout">Sign out</a>
  </div>
</div>

<div class="main">

  <!-- Quick actions -->
  <div class="cards">
    <div class="card">
      <div class="card-label">Manual Run</div>
      <div class="card-action">Trigger wages now for the current week.</div>
      <button class="btn btn-primary" onclick="triggerRun()">Run wages now</button>
    </div>
    <div class="card">
      <div class="card-label">Regenerate Month</div>
      <div class="card-action">Wipe this month's file and start fresh. Use after fixing staff.</div>
      <button class="btn btn-danger" onclick="confirmRegenerate()">Regenerate</button>
    </div>
    <div class="card">
      <div class="card-label">Download File</div>
      <div class="card-action">Download the current month's wages Excel file.</div>
      <a href="/download?key={ADMIN_PASSWORD}" class="btn btn-secondary">Download</a>
    </div>
  </div>

  <!-- Staff management -->
  <div class="section">
    <div class="section-title">Staff &amp; Rates</div>
    <p style="font-size:13px;color:#64748b;margin-bottom:20px;">
      Changes here update the master template. They take effect from next month onwards,
      or immediately if you regenerate the current month.
    </p>
    <div class="tabs">{location_tabs}</div>
    {location_panels}
  </div>

  <!-- Upload template -->
  <div class="section">
    <div class="section-title">Upload Master Template</div>
    <p style="font-size:13px;color:#64748b;margin-bottom:16px;">
      Replace the master template with a new Excel file. Staff and rates above will be overwritten.
    </p>
    <div class="upload-form">
      <input type="file" id="template-file" class="file-input" accept=".xlsx">
      <button class="btn btn-secondary" onclick="uploadTemplate()">Upload</button>
    </div>
    {"<div style='margin-top:12px'>" + files_html + "</div>" if files_html else ""}
  </div>

</div>

<div class="toast" id="toast"></div>

<script>
function showTab(loc) {{
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.getElementById('panel-' + loc).classList.add('active');
  document.getElementById('tab-' + loc).classList.add('active');
}}

function addRow(loc) {{
  const tbody = document.querySelector('#table-' + loc + ' tbody');
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input class="cell-input" name="name" placeholder="Full name"></td>
    <td><input class="cell-input rate" name="rate_mon" value="14.15" placeholder="0.00"></td>
    <td><input class="cell-input rate" name="rate_sun" value="17.98" placeholder="0.00"></td>
    <td><button class="btn-remove" onclick="removeRow(this)">×</button></td>`;
  tbody.appendChild(tr);
}}

function removeRow(btn) {{
  btn.closest('tr').remove();
}}

async function saveLocation(loc) {{
  const rows = document.querySelectorAll('#table-' + loc + ' tbody tr');
  const staff = [];
  rows.forEach(row => {{
    const name = row.querySelector('[name=name]').value.trim();
    if (name) {{
      staff.push({{
        name,
        rate_mon: row.querySelector('[name=rate_mon]').value,
        rate_sun: row.querySelector('[name=rate_sun]').value,
      }});
    }}
  }});

  const r = await fetch('/admin/save-staff', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{ location: loc, staff }})
  }});
  const data = await r.json();
  if (data.ok) {{
    toast('Saved ' + loc, false);
  }} else {{
    toast('Error: ' + data.error, true);
  }}
}}

async function triggerRun() {{
  toast('Running wages...', false);
  const r = await fetch('/trigger?key={ADMIN_PASSWORD}', {{ method: 'POST' }});
  const data = await r.json();
  toast('Done — check email', false);
}}

async function confirmRegenerate() {{
  if (!confirm('This will wipe all hours for the current month. Are you sure?')) return;
  const r = await fetch('/regenerate?key={ADMIN_PASSWORD}');
  const data = await r.json();
  toast('Month regenerated', false);
}}

async function uploadTemplate() {{
  const file = document.getElementById('template-file').files[0];
  if (!file) {{ toast('Select a file first', true); return; }}
  const fd = new FormData();
  fd.append('file', file);
  const r = await fetch('/upload-template?key={ADMIN_PASSWORD}', {{ method: 'POST', body: fd }});
  const data = await r.json();
  if (data.status === 'uploaded') {{
    toast('Template uploaded', false);
  }} else {{
    toast('Upload failed', true);
  }}
}}

function toast(msg, isError) {{
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  setTimeout(() => t.className = 'toast', 3000);
}}
</script>
</body>
</html>'''


@admin_bp.route('/admin/save-staff', methods=['POST'])
@login_required
def save_staff():
    data     = request.get_json()
    location = data.get('location')
    staff    = data.get('staff', [])

    if not location or location not in LOCATION_SHEETS:
        return jsonify({'ok': False, 'error': 'Invalid location'})

    ok, msg = save_staff_to_template(location, staff)
    return jsonify({'ok': ok, 'message': msg})
