from supabase import create_client, Client
from config import Config

_client: Client = None

def get_db() -> Client:
    global _client
    if _client is None:
        print(f"[DB] Connecting to Supabase: {Config.SUPABASE_URL[:40]}...")
        _client = create_client(Config.SUPABASE_URL, Config.SUPABASE_KEY)
        print("[DB] Client created OK")
    return _client

def fetch_one(table, match: dict):
    try:
        res = get_db().table(table).select('*').match(match).limit(1).execute()
        print(f"[DB] fetch_one {table} match={match} -> {len(res.data)} rows")
        return res.data[0] if res.data else None
    except Exception as e:
        print(f"[DB ERROR] fetch_one {table}: {e}")
        return None

def fetch_all(table, match: dict = None, order=None, limit=200):
    try:
        q = get_db().table(table).select('*')
        if match:
            q = q.match(match)
        if order:
            q = q.order(order, desc=True)
        q = q.limit(limit)
        res = q.execute()
        print(f"[DB] fetch_all {table} -> {len(res.data)} rows")
        return res.data or []
    except Exception as e:
        print(f"[DB ERROR] fetch_all {table}: {e}")
        return []

def insert_row(table, data: dict):
    try:
        res = get_db().table(table).insert(data).execute()
        print(f"[DB] insert {table} -> {res.data}")
        return res.data[0] if res.data else None
    except Exception as e:
        print(f"[DB ERROR] insert {table}: {e}")
        raise

def update_row(table, match: dict, data: dict):
    try:
        res = get_db().table(table).update(data).match(match).execute()
        print(f"[DB] update {table} match={match} -> ok")
        return res.data[0] if res.data else None
    except Exception as e:
        print(f"[DB ERROR] update {table}: {e}")
        return None

def delete_row(table, match: dict):
    try:
        get_db().table(table).delete().match(match).execute()
        print(f"[DB] delete {table} match={match} -> ok")
    except Exception as e:
        print(f"[DB ERROR] delete {table}: {e}")

def purge_old_clients(days=365):
    try:
        from datetime import datetime, timedelta
        cutoff = (datetime.utcnow() - timedelta(days=days)).date().isoformat()
        get_db().table('clients').delete().lt('created_at', cutoff).execute()
        print(f"[DB] purge_old_clients: deleted records created before {cutoff}")
    except Exception as e:
        print(f"[DB ERROR] purge_old_clients: {e}")