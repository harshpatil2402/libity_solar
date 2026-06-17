from flask import Blueprint, render_template, request, redirect, url_for, session, flash, make_response
from config import Config
from auth import superadmin_required
from db import fetch_all, fetch_one, insert_row, update_row, delete_row
from datetime import datetime
import uuid, threading
from utils.whatsapp import open_whatsapp

sa_bp = Blueprint('superadmin', __name__, url_prefix='/superadmin')

PLANS = ['full', 'docs_only']


# ── Login ─────────────────────────────────────────────────────────

@sa_bp.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('superadmin'):
        return redirect(url_for('superadmin.dashboard'))
    if request.method == 'POST':
        if (request.form.get('username') == Config.SUPERADMIN_USER and
                request.form.get('password') == Config.SUPERADMIN_PASS):
            session['superadmin'] = True
            session.permanent = True
            from db import purge_old_clients
            threading.Thread(target=purge_old_clients, daemon=True).start()
            return redirect(url_for('superadmin.dashboard'))
        flash('Invalid credentials.', 'danger')
    return render_template('superadmin/login.html')


@sa_bp.route('/logout')
def logout():
    session.pop('superadmin', None)
    return redirect(url_for('superadmin.login'))


# ── Dashboard ─────────────────────────────────────────────────────

@sa_bp.route('/')
@superadmin_required
def dashboard():
    agencies = fetch_all('agencies', order='created_at')
    total    = len(agencies)
    active   = sum(1 for a in agencies if a.get('is_active'))
    expired  = sum(1 for a in agencies if a.get('expires_at') and
                   str(a['expires_at']) < datetime.today().strftime('%Y-%m-%d'))
    return render_template('superadmin/dashboard.html',
                           agencies=agencies, total=total,
                           active=active, expired=expired,
                           today=datetime.today().strftime('%Y-%m-%d'),
                           base_url=Config.BASE_URL)


# ── Create Agency ─────────────────────────────────────────────────

@sa_bp.route('/agencies/new', methods=['GET', 'POST'])
@superadmin_required
def new_agency():
    if request.method == 'POST':
        f = request.form
        logo_b64  = _read_image_b64('logo')
        stamp_b64 = _read_image_b64('stamp')

        agency = {
            'agency_name':   f['agency_name'],
            'director_name': f.get('director_name', ''),
            'email':         f['email'].strip().lower(),
            'username':      f['username'].strip().lower(),
            'password':      f['password'].strip(),
            'contact_number':f.get('contact_number', ''),
            'agency_address':f.get('agency_address', ''),
            'plan':          f.get('plan', 'full'),
            'is_active':     True,
            'expires_at':    f.get('expires_at') or None,
            'logo_b64':      logo_b64,
            'stamp_b64':     stamp_b64,
        }
        try:
            insert_row('agencies', agency)
            flash(f"Agency '{agency['agency_name']}' created.", 'success')
            return redirect(url_for('superadmin.dashboard'))
        except Exception as e:
            flash(f'Error: {e}', 'danger')
    return render_template('superadmin/agency_form.html', agency=None, plans=PLANS, title='New Agency')


# ── Edit Agency ───────────────────────────────────────────────────

@sa_bp.route('/agencies/<agency_id>/edit', methods=['GET', 'POST'])
@superadmin_required
def edit_agency(agency_id):
    agency = fetch_one('agencies', {'id': agency_id})
    if not agency:
        flash('Agency not found.', 'danger')
        return redirect(url_for('superadmin.dashboard'))

    if request.method == 'POST':
        f = request.form
        updates = {
            'agency_name':   f['agency_name'],
            'director_name': f.get('director_name', ''),
            'email':         f['email'].strip().lower(),
            'username':      f['username'].strip().lower(),
            'contact_number':f.get('contact_number', ''),
            'agency_address':f.get('agency_address', ''),
            'plan':          f.get('plan', 'full'),
            'is_active':     f.get('is_active') == 'true',
            'expires_at':    f.get('expires_at') or None,
        }
        if f.get('password', '').strip():
            updates['password'] = f['password'].strip()

        logo_b64  = _read_image_b64('logo')
        stamp_b64 = _read_image_b64('stamp')
        if logo_b64:  updates['logo_b64']  = logo_b64
        if stamp_b64: updates['stamp_b64'] = stamp_b64

        update_row('agencies', {'id': agency_id}, updates)
        flash('Agency updated.', 'success')
        return redirect(url_for('superadmin.dashboard'))

    return render_template('superadmin/agency_form.html', agency=agency, plans=PLANS, title='Edit Agency')


# ── Toggle Active ─────────────────────────────────────────────────

@sa_bp.route('/agencies/<agency_id>/toggle', methods=['POST'])
@superadmin_required
def toggle_agency(agency_id):
    agency = fetch_one('agencies', {'id': agency_id})
    if agency:
        update_row('agencies', {'id': agency_id}, {'is_active': not agency.get('is_active', True)})
        flash('Agency status updated.', 'success')
    return redirect(url_for('superadmin.dashboard'))


# ── Delete Agency ─────────────────────────────────────────────────

@sa_bp.route('/agencies/<agency_id>/delete', methods=['POST'])
@superadmin_required
def delete_agency(agency_id):
    delete_row('agencies', {'id': agency_id})
    flash('Agency deleted.', 'success')
    return redirect(url_for('superadmin.dashboard'))


# ── View Agency Clients ───────────────────────────────────────────

@sa_bp.route('/agencies/<agency_id>/clients')
@superadmin_required
def agency_clients(agency_id):
    agency  = fetch_one('agencies', {'id': agency_id})
    clients = fetch_all('clients', {'agency_id': agency_id}, order='created_at')
    return render_template('superadmin/agency_clients.html', agency=agency, clients=clients)


# ── Export Agencies CSV ───────────────────────────────────────────

@sa_bp.route('/export-csv')
@superadmin_required
def export_csv():
    import csv, io
    agencies = fetch_all('agencies', order='created_at')
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(['Agency Name', 'Username', 'Director', 'Email', 'Contact', 'Plan', 'Active', 'Expires'])
    for a in agencies:
        w.writerow([
            a.get('agency_name', ''), a.get('username', ''), a.get('director_name', ''),
            a.get('email', ''), a.get('contact_number', ''), a.get('plan', ''),
            a.get('is_active', ''), a.get('expires_at', ''),
        ])
    resp = make_response(output.getvalue())
    resp.headers['Content-Type'] = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = 'attachment; filename=agencies.csv'
    return resp


# ── Send Credentials via WhatsApp ────────────────────────────────

@sa_bp.route('/agencies/<agency_id>/send-credentials', methods=['POST'])
@superadmin_required
def send_agency_credentials(agency_id):
    agency = fetch_one('agencies', {'id': agency_id})
    if not agency:
        flash('Agency not found.', 'danger')
        return redirect(url_for('superadmin.dashboard'))

    mobile = (agency.get('contact_number') or '').strip()
    if not mobile:
        flash(f"No contact number for {agency['agency_name']}. Edit the agency and add one first.", 'warning')
        return redirect(url_for('superadmin.dashboard'))

    login_link = f"{Config.BASE_URL}/login/{agency['username']}"
    message = (
        f"Hello {agency['agency_name']},\n\n"
        f"Your LibityInfotech Solar Portal credentials:\n"
        f"Username: {agency['username']}\n"
        f"Password: {agency['password']}\n\n"
        f"Login here:\n{login_link}\n\n"
        f"— LibityInfotech Solar Support"
    )

    def go():
        open_whatsapp(mobile, message)

    threading.Thread(target=go, daemon=True).start()
    flash(f"Opening WhatsApp for {agency['agency_name']} ({mobile})…", 'success')
    return redirect(url_for('superadmin.dashboard'))


# ── Helper ────────────────────────────────────────────────────────

def _read_image_b64(field_name: str) -> str | None:
    import base64
    file = request.files.get(field_name)
    if file and file.filename:
        return 'data:image/jpeg;base64,' + base64.b64encode(file.read()).decode()
    # Also accept raw base64 from hidden input
    b64 = request.form.get(f'{field_name}_b64', '')
    return b64 if b64 else None
