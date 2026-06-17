from functools import wraps
from flask import session, redirect, url_for, flash

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

def superadmin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('superadmin'):
            return redirect(url_for('superadmin.login'))
        return f(*args, **kwargs)
    return decorated

def agency_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        aid = session.get('agency_id')
        if not aid:
            return redirect(url_for('agency.login'))

        # Hardcoded fallback
        if aid == 'hardcoded-admin':
            kwargs['agency'] = HARDCODED_AGENCY
            return f(*args, **kwargs)

        from db import fetch_one
        agency = fetch_one('agencies', {'id': aid})
        print(f"[auth] agency_required: id={aid} -> {agency is not None}")
        if not agency:
            session.clear()
            return redirect(url_for('agency.login'))
        if not agency.get('is_active'):
            flash('Account deactivated. Contact support.', 'danger')
            session.clear()
            return redirect(url_for('agency.login'))
        # Expired agencies are allowed through — dashboard shows soft-lock overlay
        kwargs['agency'] = agency
        return f(*args, **kwargs)
    return decorated

def feature_required(feature):
    def decorator(f):
        @wraps(f)
        def decorated(*args, **kwargs):
            agency = kwargs.get('agency')
            plan = agency.get('plan', 'full') if agency else 'full'
            if feature == 'client_mgmt' and plan == 'docs_only':
                flash('Client management not enabled for your plan.', 'warning')
                return redirect(url_for('agency.dashboard'))
            return f(*args, **kwargs)
        return decorated
    return decorator