from flask import (Blueprint, render_template, request, redirect,
                   url_for, session, flash, jsonify, send_file, make_response)
from config import Config
from auth import agency_required, feature_required
from db import fetch_all, fetch_one, insert_row, update_row, delete_row
from utils.helpers import send_sms, make_token
from utils.doc_engine import jobs, jobs_lock, run_job
from datetime import datetime, timedelta, timezone
import uuid, threading, io
from utils.whatsapp import open_whatsapp, build_message

agency_bp = Blueprint('agency', __name__, url_prefix='/agency')

STATUSES = [
    'New Lead', 'Link Sent', 'Info Collected',
    'Documents Generated', 'Portal Applied', 'Sanctioned',
    'Installation Done', 'GPS Photos Received',
    'Final Submission Done', 'Completed',
]

STATUS_COLORS = {
    'New Lead':             '#406093',
    'Link Sent':            '#B05000',
    'Info Collected':       '#2D7A4A',
    'Documents Generated':  '#2050B0',
    'Portal Applied':       '#7040B0',
    'Sanctioned':           '#2D7A4A',
    'Installation Done':    '#2050B0',
    'GPS Photos Received':  '#2D7A4A',
    'Final Submission Done':'#7040B0',
    'Completed':            '#2D7A4A',
}

# ── Hardcoded fallback credentials ───────────────────────────────
HARDCODED_AGENCY = {
    'id': 'hardcoded-admin',
    'agency_name': 'Admin Agency',
    'username': 'admin',
    'password': 'admin123',
    'is_active': True,
    'plan': 'full',
    'expires_at': None,
    'director_name': 'Admin',
    'contact_number': '',
    'agency_address': '',
    'logo_b64': None,
    'stamp_b64': None,
}

# ── Login / Logout ────────────────────────────────────────────────

@agency_bp.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('agency_id'):
        return redirect(url_for('agency.dashboard'))
    err = None
    if request.method == 'POST':
        username = request.form.get('username', '').strip().lower()
        password = request.form.get('password', '')
        print(f"[LOGIN] Attempt: username={username}")

        # Hardcoded fallback
        if username == 'admin' and password == 'admin123':
            print("[LOGIN] Hardcoded admin match — logging in")
            session['agency_id'] = HARDCODED_AGENCY['id']
            session.permanent = True
            return redirect(url_for('agency.dashboard'))

        # Supabase lookup
        agency = fetch_one('agencies', {'username': username})
        print(f"[LOGIN] Supabase result: {agency}")
        if agency is None:
            err = 'Invalid username or password.'
        elif agency.get('password', '').strip() != password.strip():
            err = 'Invalid username or password.'
        elif not agency.get('is_active'):
            err = 'Account deactivated. Contact support.'
        else:
            session['agency_id'] = agency['id']
            session.permanent = True
            return redirect(url_for('agency.dashboard'))
    return render_template('agency/login.html', err=err)


@agency_bp.route('/logout')
def logout():
    session.pop('agency_id', None)
    return redirect(url_for('agency.login'))


# ── Dashboard ─────────────────────────────────────────────────────

@agency_bp.route('/')
@agency_required
def dashboard(agency):
    plan = agency.get('plan', 'full')
    if plan == 'docs_only':
        return redirect(url_for('agency.generate_page'))

    search  = request.args.get('q', '').strip()
    fstatus = request.args.get('status', '')

    clients = _get_clients(agency['id'], search, fstatus)

    # Mark unseen as seen
    unseen_ids = [c['id'] for c in clients if not c.get('seen')]
    for cid in unseen_ids:
        update_row('clients', {'id': cid}, {'seen': True})

    kanban = {s: [] for s in STATUSES}
    for c in clients:
        if c['status'] in kanban:
            kanban[c['status']].append(c)

    total     = len(clients)
    completed = sum(1 for c in clients if c['status'] == 'Completed')
    notif_count = sum(1 for c in clients if c['status'] == 'Info Collected')

    days_remaining = None
    if agency.get('expires_at'):
        try:
            exp_date = datetime.strptime(str(agency['expires_at'])[:10], '%Y-%m-%d').date()
            days_remaining = (exp_date - datetime.today().date()).days
        except Exception:
            pass

    is_expired = days_remaining is not None and days_remaining < 0

    return render_template('agency/dashboard.html',
                           agency=agency, clients=clients,
                           kanban=kanban, statuses=STATUSES,
                           status_colors=STATUS_COLORS,
                           search=search, fstatus=fstatus,
                           total=total, completed=completed,
                           notif_count=notif_count, plan=plan,
                           days_remaining=days_remaining,
                           is_expired=is_expired)


# ── New Client ────────────────────────────────────────────────────

@agency_bp.route('/clients/new', methods=['GET', 'POST'])
@agency_required
@feature_required('client_mgmt')
def new_client(agency):
    if request.method == 'POST':
        f = request.form
        client = {
            'agency_id':      agency['id'],
            'name':           f['name'],
            'mobile':         f['mobile'],
            'kw_capacity':    float(f['kw_capacity']) if f.get('kw_capacity') else None,
            'final_amount':   float(f['final_amount']) if f.get('final_amount') else None,
            'consumer_number':f.get('consumer_number', ''),
            'address':        f.get('address', ''),
            'city':           f.get('city', ''),
            'status':         'New Lead',
            'seen':           True,
        }
        row = insert_row('clients', client)
        flash(f"Client {client['name']} added.", 'success')
        return redirect(url_for('agency.client_detail', client_id=row['id']))
    return render_template('agency/new_client.html', agency=agency)


# ── Client Detail ─────────────────────────────────────────────────

@agency_bp.route('/clients/<client_id>')
@agency_required
@feature_required('client_mgmt')
def client_detail(agency, client_id):
    client = _get_client(client_id, agency['id'])
    if not client:
        flash('Client not found.', 'danger')
        return redirect(url_for('agency.dashboard'))
    submission  = fetch_one('client_submissions', {'client_id': client_id})
    gps         = fetch_all('gps_photos', {'client_id': client_id, 'status': 'submitted'})
    gps_all_pending = fetch_all('gps_photos', {'client_id': client_id, 'status': 'pending'}, order='created_at')
    gps_pending = gps_all_pending[-1] if gps_all_pending else None
    doc_jobs    = fetch_all('doc_jobs', {'client_id': client_id}, order='created_at')
    return render_template('agency/client_detail.html',
                           agency=agency, client=client,
                           submission=submission, gps_photos=gps,
                           gps_pending=gps_pending,
                           doc_jobs=doc_jobs, statuses=STATUSES,
                           status_colors=STATUS_COLORS,
                           base_url=Config.BASE_URL)


# ── Send Info Collection Link ─────────────────────────────────────

@agency_bp.route('/clients/<client_id>/send-link', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def send_collection_link(agency, client_id):
    client = _get_client(client_id, agency['id'])
    if not client:
        flash('Client not found.', 'danger')
        return redirect(url_for('agency.dashboard'))

    token   = make_token()
    expires = (datetime.now(timezone.utc) + timedelta(hours=72)).isoformat()

    insert_row('client_submissions', {
        'client_id':  client_id,
        'agency_id':  agency['id'],
        'token':      token,
        'status':     'pending',
        'expires_at': expires,
    })

    update_row('clients', {'id': client_id}, {
        'status':     'Link Sent',
        'updated_at': datetime.now(timezone.utc).isoformat(),
    })

    link = f"{Config.BASE_URL}/{agency['username']}/client/{token}"
    amount_str = f"₹{int(client['final_amount']):,}" if client.get('final_amount') else ''
    msg = (f"Hi {client['name']}, please fill your solar installation details "
           f"for {amount_str} agreement with {agency['agency_name']}. "
           f"Link: {link} (valid 72 hours)")
    send_sms(client['mobile'], msg)

    flash(f"Link sent to {client['mobile']}. Link: {link}", 'success')
    return redirect(url_for('agency.client_detail', client_id=client_id))


# ── Send GPS Photo Link ───────────────────────────────────────────

@agency_bp.route('/clients/<client_id>/send-gps-link', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def send_gps_link(agency, client_id):
    client = _get_client(client_id, agency['id'])
    token   = make_token()
    expires = (datetime.now(timezone.utc) + timedelta(hours=72)).isoformat()
    insert_row('gps_photos', {
        'client_id':  client_id,
        'agency_id':  agency['id'],
        'token':      token,
        'status':     'pending',
        'created_at': datetime.now(timezone.utc).isoformat(),
    })
    link = f"{Config.BASE_URL}/gps/{token}"
    send_sms(client['mobile'], f"Hi {client['name']}, take a GPS photo of your solar installation: {link}")
    flash(f"GPS photo link sent. Link: {link}", 'success')
    return redirect(url_for('agency.client_detail', client_id=client_id))


# ── Update Status ─────────────────────────────────────────────────

@agency_bp.route('/clients/<client_id>/status', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def update_status(agency, client_id):
    new_status = request.form.get('status')
    if new_status in STATUSES:
        upd = {'status': new_status, 'updated_at': datetime.now(timezone.utc).isoformat()}
        if new_status == 'Portal Applied':
            upd['portal_applied_at'] = datetime.now(timezone.utc).isoformat()
        update_row('clients', {'id': client_id, 'agency_id': agency['id']}, upd)
        flash(f'Status updated to "{new_status}".', 'success')
    return redirect(url_for('agency.client_detail', client_id=client_id))


# ── Delete Client ─────────────────────────────────────────────────

@agency_bp.route('/clients/<client_id>/delete', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def delete_client(agency, client_id):
    delete_row('clients', {'id': client_id, 'agency_id': agency['id']})
    flash('Client deleted.', 'success')
    return redirect(url_for('agency.dashboard'))


# ── Generate Docs page ────────────────────────────────────────────

@agency_bp.route('/generate')
@agency_required
def generate_page(agency):
    client_id  = request.args.get('client_id')
    client     = None
    submission = None
    if client_id:
        client     = _get_client(client_id, agency['id'])
        submission = fetch_one('client_submissions', {'client_id': client_id, 'status': 'submitted'})
    return render_template('agency/generate.html',
                           agency=agency, client=client, submission=submission)


# ── API: Start Generation ─────────────────────────────────────────

@agency_bp.route('/api/generate', methods=['POST'])
@agency_required
def api_generate(agency):
    fd = request.get_json()
    if not fd:
        return jsonify({'error': 'No data'}), 400

    client_id = fd.pop('client_id', None)
    jid       = str(uuid.uuid4())

    submission = {}
    if client_id:
        sub = fetch_one('client_submissions', {'client_id': client_id, 'status': 'submitted'})
        if sub:
            submission = sub

    insert_row('doc_jobs', {
        'id':              jid,
        'agency_id':       agency['id'],
        'client_id':       client_id,
        'consumer_name':   fd.get('consumer_name', ''),
        'consumer_number': fd.get('consumer_number', ''),
        'status':          'running',
    })

    with jobs_lock:
        jobs[jid] = {'logs': [], 'status': 'running', 'zip_bytes': None, 'zip_name': None}

    _agency = dict(agency)
    from flask import current_app
    _app = current_app._get_current_object()

    def go():
        with _app.app_context():
            run_job(jid, dict(fd), _agency, submission)

    threading.Thread(target=go, daemon=True).start()

    if client_id:
        update_row('clients', {'id': client_id}, {
            'status':     'Documents Generated',
            'updated_at': datetime.now(timezone.utc).isoformat(),
        })

    return jsonify({'job_id': jid})


# ── API: Job Status ───────────────────────────────────────────────

@agency_bp.route('/api/job/<jid>/status')
@agency_required
def api_job_status(agency, jid):
    since = int(request.args.get('since', 0))
    with jobs_lock:
        job = jobs.get(jid)
    if not job:
        rec = fetch_one('doc_jobs', {'id': jid, 'agency_id': agency['id']})
        if rec and rec.get('status') == 'done':
            return jsonify({'logs': [], 'total': 0, 'status': 'done',
                            'download_url': f'/agency/api/job/{jid}/download',
                            'zip_name': rec.get('zip_name')})
        return jsonify({'error': 'not found'}), 404
    dl = f'/agency/api/job/{jid}/download' if job.get('status') == 'done' else None
    return jsonify({'logs': job['logs'][since:], 'total': len(job['logs']),
                    'status': job['status'], 'download_url': dl,
                    'zip_name': job.get('zip_name')})


# ── API: Download ZIP ─────────────────────────────────────────────

@agency_bp.route('/api/job/<jid>/download')
@agency_required
def api_job_download(agency, jid):
    with jobs_lock:
        job = jobs.get(jid)
        if not job or job.get('status') != 'done' or not job.get('zip_bytes'):
            return jsonify({'error': 'Not ready'}), 404
        raw   = job['zip_bytes']
        zname = job.get('zip_name', 'documents.zip')
    return send_file(io.BytesIO(raw), mimetype='application/zip',
                     as_attachment=True, download_name=zname)


# ── Export Clients CSV ───────────────────────────────────────────

@agency_bp.route('/export-csv')
@agency_required
def export_csv(agency):
    import csv, io
    clients = _get_clients(agency['id'])
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(['Name', 'Mobile', 'Consumer No.', 'City', 'kW', 'Amount (₹)', 'Status', 'Address', 'Created'])
    for c in clients:
        w.writerow([
            c.get('name', ''), c.get('mobile', ''), c.get('consumer_number', ''),
            c.get('city', ''), c.get('kw_capacity', ''), c.get('final_amount', ''),
            c.get('status', ''), c.get('address', ''),
            str(c.get('created_at', ''))[:10],
        ])
    resp = make_response(output.getvalue())
    resp.headers['Content-Type'] = 'text/csv; charset=utf-8'
    resp.headers['Content-Disposition'] = (
        f'attachment; filename={agency["agency_name"].replace(" ","_")}_clients.csv'
    )
    return resp


# ── History ───────────────────────────────────────────────────────

@agency_bp.route('/history')
@agency_required
def history(agency):
    doc_jobs = fetch_all('doc_jobs', {'agency_id': agency['id']}, order='created_at')
    return render_template('agency/history.html', agency=agency, doc_jobs=doc_jobs)


@agency_bp.route('/history/<jid>/download')
@agency_required
def download_history(agency, jid):
    with jobs_lock:
        job = jobs.get(jid)
    if job and job.get('zip_bytes'):
        return send_file(io.BytesIO(job['zip_bytes']), mimetype='application/zip',
                         as_attachment=True, download_name=job.get('zip_name', 'documents.zip'))
    flash('File no longer available. Please regenerate.', 'warning')
    return redirect(url_for('agency.generate_page'))


# ── Manual Upload (agency uploads on client's behalf) ─────────────

@agency_bp.route('/clients/<client_id>/manual-upload', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def manual_upload(agency, client_id):
    client = _get_client(client_id, agency['id'])
    if not client:
        flash('Client not found.', 'danger')
        return redirect(url_for('agency.dashboard'))

    aadhar_b64    = _file_to_b64('aadhar')
    pan_b64       = _file_to_b64('pan')
    cheque_b64    = _file_to_b64('cheque')
    signature_b64 = _file_to_b64('signature')
    gps_b64       = _file_to_b64('gps_photo')

    has_docs = any([aadhar_b64, pan_b64, cheque_b64, signature_b64])
    has_gps  = bool(gps_b64)

    if has_docs:
        now = datetime.now(timezone.utc).isoformat()
        img = {}
        if aadhar_b64:    img['aadhar_b64']    = aadhar_b64
        if pan_b64:       img['pan_b64']       = pan_b64
        if cheque_b64:    img['cheque_b64']    = cheque_b64
        if signature_b64: img['signature_b64'] = signature_b64

        existing = fetch_one('client_submissions', {'client_id': client_id})
        if existing:
            update_row('client_submissions', {'token': existing['token']},
                       {**img, 'status': 'submitted', 'submitted_at': now})
        else:
            insert_row('client_submissions', {
                'client_id':    client_id,
                'agency_id':    agency['id'],
                'token':        make_token(),
                'status':       'submitted',
                'submitted_at': now,
                'expires_at':   (datetime.now(timezone.utc) + timedelta(hours=72)).isoformat(),
                'aadhar_b64':   aadhar_b64 or '',
                'pan_b64':      pan_b64 or '',
                'cheque_b64':   cheque_b64 or '',
                'signature_b64': signature_b64 or '',
            })

        update_row('clients', {'id': client_id}, {
            'status': 'Info Collected', 'seen': False, 'updated_at': now,
        })
        flash('Documents uploaded — status set to "Info Collected".', 'success')

    if has_gps:
        now = datetime.now(timezone.utc).isoformat()
        insert_row('gps_photos', {
            'client_id':  client_id,
            'agency_id':  agency['id'],
            'token':      make_token(),
            'status':     'submitted',
            'photo_b64':  gps_b64,
            'latitude':   None,
            'longitude':  None,
            'taken_at':   now,
            'created_at': now,
        })
        update_row('clients', {'id': client_id}, {
            'status': 'GPS Photos Received', 'updated_at': now,
        })
        if not has_docs:
            flash('GPS photo uploaded — status set to "GPS Photos Received".', 'success')

    if not has_docs and not has_gps:
        flash('No files were selected.', 'warning')

    return redirect(url_for('agency.client_detail', client_id=client_id))


# ── Send WhatsApp ─────────────────────────────────────────────────

@agency_bp.route('/clients/<client_id>/whatsapp', methods=['POST'])
@agency_required
@feature_required('client_mgmt')
def send_whatsapp(agency, client_id):
    client = _get_client(client_id, agency['id'])
    if not client:
        return jsonify({'ok': False, 'error': 'Client not found'}), 404

    data = request.get_json() or {}
    link = data.get('link', '')
    if not link:
        return jsonify({'ok': False, 'error': 'No link provided'}), 400

    message = build_message(client, link, agency['agency_name'])
    mobile  = client.get('mobile', '')

    def go():
        open_whatsapp(mobile, message)

    threading.Thread(target=go, daemon=True).start()
    return jsonify({'ok': True})


# ── Helpers ───────────────────────────────────────────────────────


def _get_client(client_id, agency_id):
    return fetch_one('clients', {'id': client_id, 'agency_id': agency_id})


def _file_to_b64(field_name: str):
    import base64
    f = request.files.get(field_name)
    if f and f.filename:
        return 'data:image/jpeg;base64,' + base64.b64encode(f.read()).decode()
    return None


def _get_clients(agency_id, search='', status=''):
    try:
        from db import get_db
        q = get_db().table('clients').select('*').eq('agency_id', agency_id)
        if status:
            q = q.eq('status', status)
        if search:
            q = q.or_(f'name.ilike.%{search}%,mobile.ilike.%{search}%')
        q = q.order('updated_at', desc=True).limit(500)
        return q.execute().data or []
    except Exception as e:
        print(f"[_get_clients ERROR] {e}")
        return []