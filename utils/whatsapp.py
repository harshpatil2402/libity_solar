import threading
import os

BASE_DIR      = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
WA_PROFILE    = os.path.join(BASE_DIR, 'wa_profile')   # persistent session folder

_lock = threading.Lock()   # one WA operation at a time


def open_whatsapp(mobile: str, message: str) -> dict:
    """
    Opens WhatsApp Web with phone + pre-filled message.
    WhatsApp Web must already be logged in (session saved in wa_profile/).
    Does NOT auto-send — user clicks Send manually.
    Returns {'ok': True} or {'ok': False, 'error': str}
    """
    with _lock:
        try:
            from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

            with sync_playwright() as p:
                browser = p.chromium.launch_persistent_context(
                    user_data_dir=WA_PROFILE,
                    headless=False,              # must be visible so user can send
                    args=['--start-maximized'],
                    no_viewport=True,
                )

                page = browser.pages[0] if browser.pages else browser.new_page()

                # Build WhatsApp Web URL with pre-filled message
                import urllib.parse
                encoded = urllib.parse.quote(message)
                url = f"https://web.whatsapp.com/send?phone=91{mobile}&text={encoded}"

                print(f"[WA] Opening: {url[:80]}...")
                page.goto(url, wait_until='domcontentloaded', timeout=30000)

                # Wait for message input box to appear (means chat loaded)
                try:
                    page.wait_for_selector(
                        'div[contenteditable="true"][data-tab="10"], '
                        'div[contenteditable="true"][data-tab="1"], '
                        'footer div[contenteditable="true"]',
                        timeout=25000
                    )
                    print("[WA] Message box ready — user can now send.")
                    # Keep browser open — user clicks Send
                    # We return immediately; browser stays open
                    return {'ok': True}

                except PWTimeout:
                    # May be QR screen or number not on WhatsApp
                    print("[WA] Timed out waiting for message box — check if WA Web is logged in.")
                    return {'ok': False, 'error': 'WhatsApp Web not ready. Please log in at web.whatsapp.com first, then retry.'}

        except Exception as e:
            print(f"[WA ERROR] {e}")
            return {'ok': False, 'error': str(e)}


def build_message(client: dict, link: str, agency_name: str) -> str:
    """Build bilingual English + Marathi message."""
    name        = client.get('name', '')
    amount      = client.get('final_amount')
    kw          = client.get('kw_capacity')
    amount_str  = f"₹{int(amount):,}" if amount else ''
    kw_str      = f"{kw} kW" if kw else ''

    english = (
        f"Dear {name},\n\n"
        f"{agency_name} has sent you a link to complete your solar installation documents.\n"
        + (f"Solar System: {kw_str}\n" if kw_str else '')
        + (f"Agreement Amount: {amount_str}\n" if amount_str else '')
        + f"\nPlease upload your Aadhaar, PAN, cancelled cheque, and signature using the link below:\n"
        f"{link}\n\n"
        f"The link is valid for 72 hours."
    )

    marathi = (
        f"\n---\n"
        f"प्रिय {name},\n\n"
        f"{agency_name} यांनी तुम्हाला सौर ऊर्जा इंस्टॉलेशनसाठी कागदपत्रे भरण्याचा लिंक पाठवला आहे.\n"
        + (f"सोलर सिस्टम: {kw_str}\n" if kw_str else '')
        + (f"करार रक्कम: {amount_str}\n" if amount_str else '')
        + f"\nखालील लिंकवर तुमचे आधार, पॅन, रद्द केलेला चेक आणि स्वाक्षरी अपलोड करा:\n"
        f"{link}\n\n"
        f"हा लिंक ७२ तासांसाठी वैध आहे."
    )

    return english + marathi