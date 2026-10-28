import os
import base64
import logging
import requests
from datetime import date, timedelta
from pathlib import Path

log = logging.getLogger(__name__)


def _send_email(to_emails: list, subject: str, html_body: str, attachments: list):
    """
    Send an HTML email with optional attachments via SendGrid.
    attachments: list of (filename, bytes_data)
    """
    api_key = os.getenv('SENDGRID_API_KEY')
    from_email = os.getenv('EMAIL_FROM', 'wages@memoryblock.com.au')

    if not api_key:
        log.warning("SENDGRID_API_KEY not set — skipping email")
        return False

    payload = {
        "personalizations": [{"to": [{"email": e} for e in to_emails]}],
        "from": {"email": from_email, "name": "MB Ireland Wages"},
        "subject": subject,
        "content": [{"type": "text/html", "value": html_body}],
        "attachments": [
            {
                "content": base64.b64encode(file_bytes).decode(),
                "type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "filename": filename,
            }
            for filename, file_bytes in attachments
        ],
    }

    r = requests.post(
        'https://api.sendgrid.com/v3/mail/send',
        json=payload,
        headers={'Authorization': f'Bearer {api_key}', 'Content-Type': 'application/json'},
    )
    if r.status_code == 202:
        log.info(f"Email sent to {', '.join(to_emails)}")
        return True
    else:
        log.error(f"SendGrid error {r.status_code}: {r.text}")
        return False



def send_wages_email(file_path: str, summary: dict, target_sunday: date = None, preliminary: bool = False):
    """Email the updated wages file with a run summary."""
    to_emails  = [e.strip() for e in os.getenv('EMAIL_TO', 'yuvi@memoryblock.com.au').split(',')]

    if target_sunday is None:
        from square_service import get_previous_week_sunday
        target_sunday = get_previous_week_sunday()

    week_start = target_sunday - timedelta(days=6)
    prelim_tag = " ⚠ PRELIMINARY — last-day shifts may be incomplete" if preliminary else ""
    subject    = f"MB Ireland Wages — {week_start.strftime('%d %b')} to {target_sunday.strftime('%d %b %Y')}{prelim_tag}"

    rows = []
    for location, data in summary.items():
        if 'error' in data:
            rows.append(f"""
            <tr style="background:#fff3cd">
              <td style="padding:8px 12px;font-weight:bold">{location}</td>
              <td colspan="3" style="padding:8px 12px;color:#856404">{data['error']}</td>
            </tr>""")
            continue

        updated   = data.get('updated', [])
        unmatched = data.get('unmatched', [])
        note      = data.get('note', '')

        if note:
            rows.append(f"""
            <tr>
              <td style="padding:8px 12px;font-weight:bold">{location}</td>
              <td colspan="3" style="padding:8px 12px;color:#6c757d">{note}</td>
            </tr>""")
            continue

        income      = data.get('income', 0.0)
        income_html = f"<div style='font-size:13px;font-weight:bold;color:#2c7a2c;margin-bottom:6px'>Income: €{income:,.2f}</div>" if income else ''

        staff_lines = ''.join(
            f"<div style='margin:2px 0;font-size:13px'>{u['name']}: "
            f"{u['weekday_hours']}h Mon-Sat, {u['sunday_hours']}h Sun</div>"
            for u in updated
        )
        unmatched_html = (
            f"<div style='color:#dc3545;font-size:12px;margin-top:4px'>"
            f"⚠ Unmatched in Square: {', '.join(unmatched)}</div>"
        ) if unmatched else ''

        added     = data.get('added', [])
        added_html = (
            f"<div style='color:#856404;font-size:12px;margin-top:4px'>"
            f"★ New staff added (default rates — check & adjust): {', '.join(added)}</div>"
        ) if added else ''

        rows.append(f"""
        <tr>
          <td style="padding:8px 12px;font-weight:bold;vertical-align:top">{location}</td>
          <td style="padding:8px 12px;vertical-align:top">{len(updated)} staff</td>
          <td style="padding:8px 12px">{income_html}{staff_lines}{unmatched_html}{added_html}</td>
        </tr>""")

    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">
      <h2 style="color:#2c3e50;border-bottom:2px solid #e9ecef;padding-bottom:10px">
        Memory Block Ireland — Weekly Wages
      </h2>
      <p style="color:#495057">
        <strong>Week:</strong> {week_start.strftime('%d %b %Y')} – {target_sunday.strftime('%d %b %Y')}
      </p>
      {"<p style='background:#2d2000;border:1px solid #7a5000;border-radius:6px;padding:10px 14px;color:#fbbf24;font-size:13px;margin-bottom:12px'>⚠ <strong>Preliminary run</strong> — this fired at 8am Melbourne time on the 1st. Ireland was still in their last evening so a few late shifts may be missing. Check Square and top up manually if needed.</p>" if preliminary else ""}
      <p>The file is attached. Check for any unmatched staff, add bonuses or adjustments (columns V, W), then process payments.</p>

      <table style="width:100%;border-collapse:collapse;margin-top:16px">
        <thead>
          <tr style="background:#343a40;color:#fff">
            <th style="padding:10px 12px;text-align:left;width:160px">Location</th>
            <th style="padding:10px 12px;text-align:left;width:80px">Staff</th>
            <th style="padding:10px 12px;text-align:left">Detail</th>
          </tr>
        </thead>
        <tbody>{''.join(rows)}</tbody>
      </table>

      <p style="color:#adb5bd;font-size:11px;margin-top:24px">
        Auto-generated by MB Ireland Wages System · Do not reply to this email
      </p>
    </div>
    """

    with open(file_path, 'rb') as f:
        file_bytes = f.read()

    _send_email(to_emails, subject, html_body, [(Path(file_path).name, file_bytes)])


def send_wages_email_cross_month(
    prev_file: str, curr_file: str,
    result_prev: dict, result_curr: dict,
    target_sunday,
    preliminary=False,
):
    """Send wages email with TWO attachments when a week crosses a month boundary."""
    import base64, calendar
    from datetime import timedelta
    from pathlib import Path

    to_emails  = [e.strip() for e in os.getenv('EMAIL_TO', 'yuvi@memoryblock.com.au').split(',')]

    week_sunday = target_sunday
    week_monday = week_sunday - timedelta(days=6)
    week_saturday = week_monday + timedelta(days=5)

    subject = (
        f"MB Ireland Wages — {week_monday.strftime('%d %b')} to "
        f"{week_saturday.strftime('%d %b')} + Sun {week_sunday.strftime('%d %b %Y')} "
        f"[cross-month]"
    )

    def _build_rows(result, title):
        rows = []
        for location, data in result.items():
            if 'error' in data:
                rows.append(f"<tr style='background:#fff3cd'><td style='padding:8px 12px;font-weight:bold'>{location}</td>"
                            f"<td colspan='2' style='padding:8px 12px;color:#856404'>{data['error']}</td></tr>")
                continue
            updated   = data.get('updated', [])
            unmatched = data.get('unmatched', [])
            added     = data.get('added', [])
            income    = data.get('income', 0.0)
            note      = data.get('note', '')
            if note:
                rows.append(f"<tr><td style='padding:8px 12px;font-weight:bold'>{location}</td>"
                            f"<td colspan='2' style='padding:8px 12px;color:#6c757d'>{note}</td></tr>")
                continue
            income_html  = f"<div style='font-size:13px;font-weight:bold;color:#2c7a2c;margin-bottom:4px'>Income: €{income:,.2f}</div>" if income else ''
            staff_lines  = ''.join(f"<div style='font-size:13px'>{u['name']}: {u['weekday_hours']}h</div>" for u in updated)
            added_html   = f"<div style='color:#856404;font-size:12px'>★ New: {', '.join(added)}</div>" if added else ''
            unmatched_html = f"<div style='color:#dc3545;font-size:12px'>⚠ Unmatched: {', '.join(unmatched)}</div>" if unmatched else ''
            rows.append(f"<tr><td style='padding:8px 12px;font-weight:bold;vertical-align:top'>{location}</td>"
                        f"<td style='padding:8px 12px;vertical-align:top'>{len(updated)} staff</td>"
                        f"<td style='padding:8px 12px'>{income_html}{staff_lines}{added_html}{unmatched_html}</td></tr>")
        return rows

    prev_rows = _build_rows(result_prev, 'prev')
    curr_rows = _build_rows(result_curr, 'curr')

    from pathlib import Path as P2
    prev_month_name = P2(prev_file).stem.split('_')[2] + ' ' + P2(prev_file).stem.split('_')[3]
    curr_month_name = P2(curr_file).stem.split('_')[2] + ' ' + P2(curr_file).stem.split('_')[3]

    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">
      <h2 style="color:#2c3e50;border-bottom:2px solid #e9ecef;padding-bottom:10px">
        Memory Block Ireland — Cross-Month Wages Week
      </h2>
      <p><strong>Week:</strong> {week_monday.strftime('%d %b %Y')} – {week_saturday.strftime('%d %b %Y')} + Sun {week_sunday.strftime('%d %b %Y')}</p>
      <p>This week spans two months. Two files are attached — one for each month.</p>

      <h3 style="margin-top:20px">{prev_month_name} portion ({week_monday.strftime('%d %b')} – {(week_saturday - timedelta(days=1) if week_saturday.month != week_monday.month else week_saturday).strftime('%d %b')})</h3>
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#343a40;color:#fff">
          <th style="padding:8px 12px;text-align:left">Location</th>
          <th style="padding:8px 12px;text-align:left">Staff</th>
          <th style="padding:8px 12px;text-align:left">Detail</th>
        </tr></thead>
        <tbody>{''.join(prev_rows)}</tbody>
      </table>

      <h3 style="margin-top:20px">{curr_month_name} portion ({week_saturday.strftime('%d %b')} + Sun {week_sunday.strftime('%d %b')})</h3>
      <table style="width:100%;border-collapse:collapse">
        <thead><tr style="background:#343a40;color:#fff">
          <th style="padding:8px 12px;text-align:left">Location</th>
          <th style="padding:8px 12px;text-align:left">Staff</th>
          <th style="padding:8px 12px;text-align:left">Detail</th>
        </tr></thead>
        <tbody>{''.join(curr_rows)}</tbody>
      </table>

      <p style="color:#adb5bd;font-size:11px;margin-top:24px">
        Auto-generated by MB Ireland Wages System
      </p>
    </div>
    """

    attachments = []
    for fp in [prev_file, curr_file]:
        with open(fp, 'rb') as f:
            attachments.append((Path(fp).name, f.read()))

    _send_email(to_emails, subject, html_body, attachments)


def send_catchup_email(file_paths: list, statuses: list, up_to):
    """
    One email after a catch-up run: week-by-week status plus the refreshed Excel file(s).
    statuses: [{'monday', 'sunday', 'partial', 'ok', 'error'}, ...]
    """
    to_emails = [e.strip() for e in os.getenv('EMAIL_TO', 'yuvi@memoryblock.com.au').split(',')]

    failed  = [s for s in statuses if not s['ok']]
    partial = [s for s in statuses if s['partial'] and s['ok']]
    first   = statuses[0]['monday']

    subject = f"MB Ireland Wages: catch-up {first.strftime('%d %b')} to {up_to.strftime('%d %b %Y')}"
    if failed:
        subject += f" ({len(failed)} week(s) FAILED)"

    rows = ''
    for s in statuses:
        label = f"{s['monday'].strftime('%d %b')} to {s['sunday'].strftime('%d %b')}"
        if not s['ok']:
            status = f"<span style='color:#dc3545;font-weight:bold'>FAILED</span> <span style='color:#6c757d;font-size:12px'>{(s['error'] or '')[:160]}</span>"
        elif s['partial']:
            status = f"<span style='color:#b45309'>Done, up to {up_to.strftime('%d %b')} only</span>"
        else:
            status = "<span style='color:#2c7a2c'>Done</span>"
        rows += f"<tr><td style='padding:8px 12px'>{label}</td><td style='padding:8px 12px'>{status}</td></tr>"

    failed_html = ''
    if failed:
        failed_html = (
            "<p style='background:#fdecea;border:1px solid #f5c2c7;border-radius:6px;padding:10px 14px;"
            "color:#842029;font-size:13px'>"
            f"{len(failed)} week(s) could not be pulled from Square, so those columns are missing from the "
            "attached file. Run the catch-up again once Square is responding.</p>")

    partial_html = ''
    if partial:
        partial_html = (
            "<p style='background:#fff8e1;border:1px solid #ffe08a;border-radius:6px;padding:10px 14px;"
            "color:#7a5000;font-size:13px'>"
            f"The last week is only counted up to {up_to.strftime('%d %b')}. "
            "It will be completed by the normal Monday run.</p>")

    html_body = f"""
    <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">
      <h2 style="color:#2c3e50;border-bottom:2px solid #e9ecef;padding-bottom:10px">
        Memory Block Ireland: Wages catch-up
      </h2>
      <p>Every week from {first.strftime('%d %b %Y')} up to {up_to.strftime('%d %b %Y')} was rebuilt
         from Square. The refreshed file(s) are attached.</p>
      {failed_html}{partial_html}
      <table style="width:100%;border-collapse:collapse;margin-top:12px">
        <thead><tr style="background:#343a40;color:#fff">
          <th style="padding:8px 12px;text-align:left">Week</th>
          <th style="padding:8px 12px;text-align:left">Status</th>
        </tr></thead>
        <tbody>{rows}</tbody>
      </table>
      <p style="color:#adb5bd;font-size:11px;margin-top:24px">
        Auto-generated by MB Ireland Wages System. Do not reply to this email.
      </p>
    </div>
    """

    attachments = []
    for fp in file_paths:
        with open(fp, 'rb') as f:
            attachments.append((Path(fp).name, f.read()))

    return _send_email(to_emails, subject, html_body, attachments)
