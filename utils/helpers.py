import requests, secrets, os
from config import Config


def send_sms(mobile: str, message: str) -> bool:
    if not Config.FAST2SMS_KEY:
        print(f"[SMS DEV] → {mobile}: {message}")
        return True
    try:
        r = requests.post(
            'https://www.fast2sms.com/dev/bulkV2',
            headers={'authorization': Config.FAST2SMS_KEY},
            data={'route': 'q', 'message': message,
                  'language': 'english', 'flash': 0, 'numbers': mobile},
            timeout=10
        )
        return r.json().get('return', False)
    except Exception as e:
        print(f"[SMS ERROR] {e}")
        return False


def make_token(n=32) -> str:
    return secrets.token_urlsafe(n)
