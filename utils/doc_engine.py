import os, re, io, base64, zipfile, uuid, threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from PIL import Image

BASE_DIR      = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
HTML_DOCS_DIR = os.path.join(BASE_DIR, 'input_docs')

_HTML_CACHE: dict = {}
jobs: dict = {}
jobs_lock = threading.Lock()


def preload_templates():
    names = [
        "commissioning_report.html",
        "meter_testing.html",
        "model_agreement.html",
        "net_metering_agreement.html",
        "work_completion_report.html",
    ]
    n = 0
    for name in names:
        path = os.path.join(HTML_DOCS_DIR, name)
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                _HTML_CACHE[name] = f.read()
            n += 1
    print(f"[DocEngine] Preloaded {n}/{len(names)} templates.")


def job_log(jid, msg, error=False):
    with jobs_lock:
        if jid in jobs:
            jobs[jid]['logs'].append({'msg': msg, 'error': error})


def b64_to_data_uri(b64: str, mime='image/jpeg') -> str:
    if not b64:
        return ''
    try:
        raw = b64.split(',')[1] if ',' in b64 else b64
        img = Image.open(io.BytesIO(base64.b64decode(raw))).convert('RGB')
        out = io.BytesIO()
        img.save(out, format='JPEG', quality=90)
        enc = base64.b64encode(out.getvalue()).decode()
        return f"data:image/jpeg;base64,{enc}"
    except Exception as e:
        print(f"[b64_to_uri] {e}")
        return ''


def file_to_data_uri(path: str, mime: str) -> str:
    if not path or not os.path.exists(path):
        return ''
    with open(path, 'rb') as f:
        return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"


def fill_template(html: str, ctx: dict) -> str:
    for k, v in ctx.items():
        if v is not None:
            html = html.replace('{{' + k + '}}', str(v))
    html = re.sub(r'\{\{[^}]+\}\}', '', html)
    html = html.replace(' class="highlight"', '').replace(" class='highlight'", '')
    return html


def render_pdf(fname: str, ctx: dict, jid: str) -> bytes | None:
    html_src = _HTML_CACHE.get(fname)
    if not html_src:
        path = os.path.join(HTML_DOCS_DIR, fname)
        if not os.path.exists(path):
            job_log(jid, f"Missing template: {fname}", error=True)
            return None
        with open(path, 'r', encoding='utf-8') as f:
            html_src = f.read()
    filled = fill_template(html_src, ctx)
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page()
            page.set_content(filled, wait_until='networkidle')
            pdf = page.pdf(print_background=True)
            browser.close()
        return pdf
    except Exception:
        try:
            from weasyprint import HTML as WP
            return WP(string=filled, base_url=HTML_DOCS_DIR).write_pdf()
        except Exception as e:
            job_log(jid, f"PDF render failed ({fname}): {e}", error=True)
            return None


def run_job(jid: str, fd: dict, agency: dict, submission: dict):
    """Main generation function — runs in background thread."""
    def log(m, e=False): job_log(jid, m, e)
    try:
        log("Loading assets ...")

        # Logos from disk
        maha_uri   = file_to_data_uri(os.path.join(BASE_DIR, 'static', 'images', 'mahavitaran_logo.png'), 'image/png')
        libity_uri = file_to_data_uri(os.path.join(BASE_DIR, 'static', 'images', 'libitylogo.png'), 'image/png')

        # Agency logo/stamp from Supabase base64
        agency_logo_uri  = b64_to_data_uri(agency.get('logo_b64', ''))  or libity_uri
        agency_stamp_uri = b64_to_data_uri(agency.get('stamp_b64', ''))

        # Client images from submission
        log("Processing client images ...")
        sig_uri    = b64_to_data_uri(submission.get('signature_b64', ''))
        aadhar_uri = b64_to_data_uri(submission.get('aadhar_b64', ''))

        inv = fd.get('inverter_make_and_model', '')
        inv_parts = inv.split(' ', 1) if inv else ['', '']

        ctx = {
            # Consumer
            'consumer_name':             fd.get('consumer_name', ''),
            'consumer_number':           fd.get('consumer_number', ''),
            'consumer_contact_number':   fd.get('consumer_contact_number', ''),
            'consumer_email':            fd.get('consumer_email', ''),
            'consumer_address':          fd.get('consumer_address', ''),
            'consumer_aadhar_num':       fd.get('consumer_aadhar_num', ''),
            'city':                      fd.get('city', ''),
            # Grid
            'discom_division':           fd.get('discom_division', ''),
            'licensee_name':             fd.get('licensee_name', ''),
            'sanction_number':           fd.get('sanction_number', ''),
            'sanction_capacity_kw':      fd.get('sanction_capacity_kw', ''),
            'system_capacity_kw':        fd.get('system_capacity_kw', ''),
            'agreement_solar_price':     fd.get('agreement_solar_price', ''),
            # Modules
            'module_make':               fd.get('module_make', ''),
            'almm_model_number':         fd.get('almm_model_number', ''),
            'module_efficiency':         fd.get('module_efficiency', ''),
            'module_capacity_wp':        fd.get('module_capacity_wp', ''),
            'num_pv_modules':            fd.get('num_pv_modules', ''),
            'total_module_capacity_kwp': fd.get('total_module_capacity_kwp', ''),
            # Inverter
            'inverter_make_and_model':   inv,
            'inverter_make':             inv_parts[0],
            'inverter_model':            inv_parts[1] if len(inv_parts) > 1 else inv,
            'inverter_capacity_kw':      fd.get('inverter_capacity_kw', ''),
            'inverter_rating_text':      fd.get('inverter_rating_text', ''),
            # Dates
            'agreement_date':            fd.get('agreement_date', ''),
            'annexure_agreement_date':   fd.get('annexure_agreement_date', ''),
            'installation_date':         fd.get('installation_date', ''),
            'meter_testing_date':        fd.get('meter_testing_date', ''),
            'performance_check_date':    fd.get('performance_check_date', ''),
            'today_date':                datetime.now().strftime('%d-%m-%Y'),
            # Agency
            'agency_name':               agency.get('agency_name', ''),
            'agency_address':            agency.get('agency_address', ''),
            'agency_contact':            agency.get('contact_number', ''),
            'agency_director':           agency.get('director_name', ''),
            # Images — inline HTML img tags
            'mahavitaran_logo':          f'<img src="{maha_uri}" style="width:100%;height:100%;object-fit:contain;">' if maha_uri else '',
            # meter_testing uses <img src="{{agency_logo}}"> so inject URI only
            'agency_logo':               agency_logo_uri if agency_logo_uri else '',
            # wcr and others that need full tag
            'agency_logo_html':          f'<img src="{agency_logo_uri}" style="width:100%;height:100%;object-fit:contain;">' if agency_logo_uri else '',
            'consumer_signature_image':  f'<img src="{sig_uri}" style="max-width:100%;max-height:100%;width:auto;height:auto;display:block;">' if sig_uri else '',
            'consumer_aadhar_image':     f'<img src="{aadhar_uri}" style="width:100%;height:100%;object-fit:contain;">' if aadhar_uri else '',
            'agency_stamp_image':        f'<img src="{agency_stamp_uri}" style="width:100%;height:100%;object-fit:contain;">' if agency_stamp_uri else '',
        }

        doc_files = [
            ("commissioning_report.html",  "1_Commissioning_Report"),
            ("meter_testing.html",          "2_Meter_Testing"),
            ("model_agreement.html",        "3_Model_Agreement"),
            ("net_metering_agreement.html", "4_Net_Metering"),
            ("work_completion_report.html", "5_Work_Completion"),
        ]

        log("Generating PDFs in parallel ...")
        pdf_results = [None] * len(doc_files)

        def render_one(idx, fname, oname):
            log(f"Rendering {oname} ...")
            pdf_results[idx] = render_pdf(fname, ctx, jid)
            if pdf_results[idx]:
                log(f"Done: {oname}.pdf")

        with ThreadPoolExecutor(max_workers=5) as ex:
            futs = {ex.submit(render_one, i, fn, on): i for i, (fn, on) in enumerate(doc_files)}
            for ft in as_completed(futs): ft.result()

        log("Building ZIP ...")
        cn    = fd.get('consumer_name', 'Client').replace(' ', '_')
        cno   = fd.get('consumer_number', '0000')
        zname = f"{cn}_{cno}_{datetime.now().strftime('%d-%m-%Y_%H%M%S')}.zip"
        zbuf  = io.BytesIO()
        with zipfile.ZipFile(zbuf, 'w', zipfile.ZIP_STORED) as zf:
            for i, (_, oname) in enumerate(doc_files):
                if pdf_results[i]:
                    zf.writestr(f"{oname}.pdf", pdf_results[i])
        zbuf.seek(0)

        with jobs_lock:
            jobs[jid]['status']    = 'done'
            jobs[jid]['zip_name']  = zname
            jobs[jid]['zip_bytes'] = zbuf.getvalue()
        log("ZIP ready — downloading.")

        # Update Supabase doc_job
        try:
            from db import update_row
            update_row('doc_jobs', {'id': jid}, {'status': 'done', 'zip_name': zname})
        except Exception: pass

    except Exception as e:
        job_log(jid, f"Fatal error: {e}", error=True)
        with jobs_lock:
            if jid in jobs:
                jobs[jid]['status'] = 'error'
        try:
            from db import update_row
            update_row('doc_jobs', {'id': jid}, {'status': 'error'})
        except Exception: pass