from flask import Blueprint, render_template, request, jsonify, redirect, url_for
from db import fetch_one, update_row
from datetime import datetime, timezone
import re

public_bp = Blueprint('public', __name__)


# ── Agency login slug: /login/<username> pre-fills the login form ──

@public_bp.route('/login/<agency_username>', methods=['GET'])
def agency_login_slug(agency_username):
    return render_template('agency/login.html', prefilled_username=agency_username, err=None)


# ── Legacy redirect: old /collect/<token> → new slug URL ─────────

@public_bp.route('/collect/<token>', methods=['GET'])
def collect_legacy(token):
    sub = fetch_one('client_submissions', {'token': token})
    if not sub:
        return render_template('public/invalid.html', reason="Link not found.")
    agency = fetch_one('agencies', {'id': sub['agency_id']})
    if not agency:
        return render_template('public/invalid.html', reason="Agency not found.")
    return redirect(f"/{agency['username']}/client/{token}", code=301)


# ── Client Info Collection ────────────────────────────────────────

@public_bp.route('/<agency_username>/client/<token>', methods=['GET', 'POST'])
def collect(agency_username, token):
    sub = fetch_one('client_submissions', {'token': token})
    if not sub:
        return render_template('public/invalid.html', reason="Link not found.")

    # Check expiry (still enforce even on re-submissions)
    exp = sub.get('expires_at')
    if exp:
        exp_str = re.sub(r'\.\d+', lambda m: (m.group() + '000000')[:7], str(exp).replace('Z', '+00:00'))
        exp_dt = datetime.fromisoformat(exp_str)
        if datetime.now(timezone.utc) > exp_dt:
            update_row('client_submissions', {'token': token}, {'status': 'expired'})
            return render_template('public/invalid.html', reason="This link has expired.")

    # Fetch client + agency for display
    client = fetch_one('clients', {'id': sub['client_id']})
    agency = fetch_one('agencies', {'id': sub['agency_id']})

    if request.method == 'POST':
        data = request.get_json()
        if not data:
            return jsonify({'error': 'No data'}), 400

        # Validate required images
        if not data.get('signature_b64'):
            return jsonify({'error': 'Signature is required'}), 400

        updates = {
            'status':       'submitted',
            'submitted_at': datetime.now(timezone.utc).isoformat(),
        }
        # Only overwrite image fields if new data was provided; keeps existing if empty
        for field in ['aadhar_b64', 'pan_b64', 'cheque_b64', 'signature_b64']:
            val = data.get(field, '')
            if val:
                updates[field] = val
        update_row('client_submissions', {'token': token}, updates)

        # Update client status
        update_row('clients', {'id': sub['client_id']}, {
            'status':     'Info Collected',
            'seen':       False,   # triggers notification badge
            'updated_at': datetime.now(timezone.utc).isoformat(),
        })

        return jsonify({'success': True})

    return render_template('public/collect.html',
                           sub=sub, client=client, agency=agency)


# ── GPS Photo ─────────────────────────────────────────────────────

@public_bp.route('/gps/<token>', methods=['GET', 'POST'])
def gps_photo(token):
    rec = fetch_one('gps_photos', {'token': token})
    if not rec:
        return render_template('public/invalid.html', reason="Link not found.")
    if rec.get('status') == 'submitted':
        return render_template('public/already_submitted.html')

    client = fetch_one('clients', {'id': rec['client_id']})
    agency = fetch_one('agencies', {'id': rec['agency_id']})

    if request.method == 'POST':
        data = request.get_json()
        if not data or not data.get('photo_b64'):
            return jsonify({'error': 'No photo'}), 400

        update_row('gps_photos', {'token': token}, {
            'status':    'submitted',
            'photo_b64': data['photo_b64'],
            'latitude':  data.get('lat'),
            'longitude': data.get('lng'),
            'taken_at':  datetime.now(timezone.utc).isoformat(),
        })
        update_row('clients', {'id': rec['client_id']}, {
            'status':     'GPS Photos Received',
            'updated_at': datetime.now(timezone.utc).isoformat(),
        })
        return jsonify({'success': True})

    return render_template('public/gps.html', rec=rec, client=client, agency=agency)
