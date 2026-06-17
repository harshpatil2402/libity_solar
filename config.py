import os
from dotenv import load_dotenv
load_dotenv()

class Config:
    SECRET_KEY        = os.environ.get('SECRET_KEY', 'dev-secret-change-in-prod')
    SUPABASE_URL      = os.environ.get('SUPABASE_URL', '')
    SUPABASE_KEY      = os.environ.get('SUPABASE_KEY', '')
    SUPERADMIN_USER   = os.environ.get('SUPERADMIN_USER', 'superadmin')
    SUPERADMIN_PASS   = os.environ.get('SUPERADMIN_PASS', 'changeme')
    FAST2SMS_KEY      = os.environ.get('FAST2SMS_KEY', '')
    BASE_URL          = os.environ.get('BASE_URL', 'http://localhost:5000')
    MAX_CONTENT_LENGTH = 50 * 1024 * 1024
