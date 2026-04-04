"""
Auth Route — syncs session across all tsifl add-ins (Excel, Word, PowerPoint, etc.)
Each Office add-in has isolated localStorage, so Supabase sessions don't share.
This module stores the session centrally so any add-in can restore it.

Uses in-memory cache + Supabase persistence.
In-memory is the fast path; Supabase survives container restarts on Railway.

Required Supabase table:
  CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY DEFAULT 'current',
    access_token TEXT,
    refresh_token TEXT,
    user_id TEXT,
    email TEXT,
    last_active TIMESTAMPTZ DEFAULT now(),
    updated_at TIMESTAMPTZ DEFAULT now()
  );

  -- Allow the backend (anon key) to read/write sessions
  ALTER TABLE sessions ENABLE ROW LEVEL SECURITY;
  CREATE POLICY "Allow all access to sessions" ON sessions FOR ALL USING (true) WITH CHECK (true);
"""

import json
import os
import time
import logging
from datetime import datetime
from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from pathlib import Path

router = APIRouter()
logger = logging.getLogger("tsifl.auth")

# In-memory session cache — fast path
_session_store = {}
_user_store = {}

# Login attempt rate limiter (Improvement 6)
_login_attempts: dict = {}  # {email: {"count": int, "first_attempt": float}}

# Filesystem backup — use Railway persistent volume if available, fallback to /tmp
# Railway volume mounted at /data survives redeploys
_DATA_DIR = Path("/data")
if _DATA_DIR.exists() and _DATA_DIR.is_dir():
    SESSION_FILE = _DATA_DIR / ".tsifl_session.json"
    logger.info("[auth] Using persistent volume at /data for session storage")
else:
    SESSION_FILE = Path("/tmp/.tsifl_session.json")
    logger.info("[auth] No persistent volume — using /tmp (sessions won't survive redeploys)")

# Track whether Supabase sessions table exists
_supabase_table_ok = False


def _save_to_file(data: dict):
    """Backup session to filesystem (survives process restarts within same deploy)."""
    try:
        SESSION_FILE.write_text(json.dumps(data))
    except Exception as e:
        logger.warning(f"[auth] File backup failed: {e}")


def _load_from_file() -> dict | None:
    """Restore session from filesystem backup."""
    try:
        if SESSION_FILE.exists():
            text = SESSION_FILE.read_text().strip()
            if text.startswith("{"):
                data = json.loads(text)
                if data.get("access_token"):
                    return data
    except Exception as e:
        logger.warning(f"[auth] File restore failed: {e}")
    return None


# Supabase client (lazy loaded)
_sb = None


def _get_supabase():
    global _sb
    if _sb is not None:
        return _sb
    url = os.getenv("SUPABASE_URL", "")
    key = os.getenv("SUPABASE_KEY", "")
    if not url or not key or url == "your_supabase_project_url":
        return None
    try:
        from supabase import create_client
        _sb = create_client(url, key)
        return _sb
    except Exception as e:
        logger.error(f"[auth] Supabase client creation failed: {e}")
        return None


async def ensure_sessions_table():
    """Try to create the sessions table if it doesn't exist.
    Uses httpx to call Supabase's SQL endpoint with the service role key."""
    global _supabase_table_ok

    # First, check if the table already exists
    client = _get_supabase()
    if not client:
        logger.warning("[auth] No Supabase client — sessions won't persist across deploys")
        return False

    try:
        result = client.table("sessions").select("id").limit(1).execute()
        _supabase_table_ok = True
        logger.info("[auth] Sessions table exists and is accessible")
        return True
    except Exception as e:
        logger.warning(f"[auth] Sessions table not found: {e}")

    # Table doesn't exist — try to create it using the service role key
    service_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY", "")
    supabase_url = os.getenv("SUPABASE_URL", "")

    if not service_key:
        logger.error(
            "[auth] CRITICAL: Sessions table does not exist in Supabase and no "
            "SUPABASE_SERVICE_ROLE_KEY is set. Sessions will NOT persist across Railway deploys.\n"
            "Fix options:\n"
            "  1. Go to Supabase Dashboard > SQL Editor and run:\n"
            "     CREATE TABLE IF NOT EXISTS sessions (\n"
            "       id TEXT PRIMARY KEY DEFAULT 'current',\n"
            "       access_token TEXT,\n"
            "       refresh_token TEXT,\n"
            "       user_id TEXT,\n"
            "       email TEXT,\n"
            "       last_active TIMESTAMPTZ DEFAULT now(),\n"
            "       updated_at TIMESTAMPTZ DEFAULT now()\n"
            "     );\n"
            "     ALTER TABLE sessions ENABLE ROW LEVEL SECURITY;\n"
            "     CREATE POLICY \"Allow all access\" ON sessions FOR ALL USING (true) WITH CHECK (true);\n"
            "  2. OR set SUPABASE_SERVICE_ROLE_KEY in Railway env vars and restart"
        )
        return False

    # Try to create via Supabase SQL endpoints
    create_sql = """
    CREATE TABLE IF NOT EXISTS sessions (
        id TEXT PRIMARY KEY DEFAULT 'current',
        access_token TEXT,
        refresh_token TEXT,
        user_id TEXT,
        email TEXT,
        last_active TIMESTAMPTZ DEFAULT now(),
        updated_at TIMESTAMPTZ DEFAULT now()
    );
    ALTER TABLE sessions ENABLE ROW LEVEL SECURITY;
    DO $$
    BEGIN
        IF NOT EXISTS (
            SELECT 1 FROM pg_policies WHERE tablename = 'sessions' AND policyname = 'Allow all access'
        ) THEN
            CREATE POLICY "Allow all access" ON sessions FOR ALL USING (true) WITH CHECK (true);
        END IF;
    END
    $$;
    """

    import httpx

    # Extract project ref from URL (e.g., "dvynmzeyttwlmvunicqz" from "https://dvynmzeyttwlmvunicqz.supabase.co")
    project_ref = supabase_url.replace("https://", "").split(".")[0]

    # Approach 1: Try the Supabase pg-meta SQL endpoint
    endpoints_to_try = [
        f"https://{project_ref}.supabase.co/rest/v1/rpc/exec_sql",
        f"https://{project_ref}.supabase.co/pg/query",
    ]

    headers = {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Content-Type": "application/json",
    }

    async with httpx.AsyncClient(timeout=15) as http:
        # Try pg/query endpoint (Supabase pg-meta)
        try:
            resp = await http.post(
                f"https://{project_ref}.supabase.co/pg/query",
                headers=headers,
                json={"query": create_sql},
            )
            if resp.status_code < 300:
                logger.info("[auth] Sessions table created via pg/query endpoint")
                _supabase_table_ok = True
                return True
            else:
                logger.warning(f"[auth] pg/query endpoint returned {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            logger.warning(f"[auth] pg/query endpoint failed: {e}")

        # Try the RPC endpoint (if exec_sql function exists)
        try:
            resp = await http.post(
                f"https://{project_ref}.supabase.co/rest/v1/rpc/exec_sql",
                headers=headers,
                json={"sql": create_sql},
            )
            if resp.status_code < 300:
                logger.info("[auth] Sessions table created via rpc/exec_sql")
                _supabase_table_ok = True
                return True
            else:
                logger.warning(f"[auth] rpc/exec_sql returned {resp.status_code}: {resp.text[:200]}")
        except Exception as e:
            logger.warning(f"[auth] rpc/exec_sql failed: {e}")

        # Try using the Supabase Management API
        try:
            mgmt_resp = await http.post(
                f"https://api.supabase.com/v1/projects/{project_ref}/database/query",
                headers={
                    "Authorization": f"Bearer {service_key}",
                    "Content-Type": "application/json",
                },
                json={"query": create_sql},
            )
            if mgmt_resp.status_code < 300:
                logger.info("[auth] Sessions table created via Management API")
                _supabase_table_ok = True
                return True
            else:
                logger.warning(f"[auth] Management API returned {mgmt_resp.status_code}: {mgmt_resp.text[:200]}")
        except Exception as e:
            logger.warning(f"[auth] Management API failed: {e}")

    # Try direct PostgreSQL connection if DATABASE_URL is set
    db_url = os.getenv("DATABASE_URL", "")
    if db_url:
        try:
            import asyncio
            import subprocess
            # Use psql if available, or psycopg2 if installed
            try:
                import psycopg2
                conn = psycopg2.connect(db_url)
                cur = conn.cursor()
                cur.execute(create_sql)
                conn.commit()
                cur.close()
                conn.close()
                logger.info("[auth] Sessions table created via direct PostgreSQL connection")
                _supabase_table_ok = True
                return True
            except ImportError:
                logger.warning("[auth] psycopg2 not installed — cannot create table via direct connection")
            except Exception as e:
                logger.warning(f"[auth] Direct PostgreSQL connection failed: {e}")
        except Exception as e:
            logger.warning(f"[auth] DATABASE_URL connection failed: {e}")

    logger.error(
        "[auth] Could not auto-create sessions table. Please create it manually in "
        "Supabase Dashboard > SQL Editor. See startup logs for the SQL."
    )
    return False


def _save_to_supabase(data: dict):
    """Persist session to Supabase sessions table."""
    global _supabase_table_ok
    client = _get_supabase()
    if not client:
        return
    try:
        row = {
            "id": "current",
            "access_token": data.get("access_token", ""),
            "refresh_token": data.get("refresh_token", ""),
            "user_id": data.get("user_id", ""),
            "email": data.get("email", ""),
            "last_active": datetime.utcnow().isoformat(),
        }
        client.table("sessions").upsert(row).execute()
        _supabase_table_ok = True
    except Exception as e:
        _supabase_table_ok = False
        logger.warning(f"[auth] Supabase save failed: {e}")


def _load_from_supabase() -> dict | None:
    """Load session from Supabase sessions table."""
    client = _get_supabase()
    if not client:
        return None
    try:
        result = client.table("sessions").select("*").eq("id", "current").execute()
        if result.data and result.data[0].get("access_token"):
            return result.data[0]
    except Exception as e:
        logger.warning(f"[auth] Supabase load failed: {e}")
    return None


def _clear_supabase():
    """Delete session from Supabase."""
    client = _get_supabase()
    if not client:
        return
    try:
        client.table("sessions").delete().eq("id", "current").execute()
    except Exception as e:
        logger.warning(f"[auth] Supabase clear failed: {e}")


class UserConfig(BaseModel):
    user_id: str
    email: str = ""


class SessionConfig(BaseModel):
    access_token: str
    refresh_token: str
    user_id: str = ""
    email: str = ""


@router.post("/set-user")
async def set_user(config: UserConfig):
    """Store user ID centrally."""
    _user_store["current"] = config.user_id
    return {"status": "ok", "user_id": config.user_id}


@router.get("/current-user")
async def current_user():
    """Read the currently saved user ID."""
    uid = _user_store.get("current")
    if not uid:
        # Try from Supabase session
        session = _load_from_supabase()
        if session and session.get("user_id"):
            uid = session["user_id"]
            _user_store["current"] = uid
    return {"user_id": uid}


@router.post("/set-session")
async def set_session(config: SessionConfig):
    """Store Supabase session tokens so all add-ins can share the login."""
    data = {
        "access_token": config.access_token,
        "refresh_token": config.refresh_token,
        "user_id": config.user_id,
        "email": config.email,
    }
    # Store in memory (cache)
    _session_store["current"] = data
    # Persist to filesystem (survives process restarts within deploy)
    _save_to_file(data)
    # Persist to Supabase (survives Railway redeploys — if table exists)
    _save_to_supabase(data)
    # Keep user store in sync
    if config.user_id:
        _user_store["current"] = config.user_id
    logger.info(f"[auth] Session stored for {config.email or 'unknown'} (supabase_ok={_supabase_table_ok})")
    return {"status": "ok", "persisted_to_supabase": _supabase_table_ok}


@router.get("/get-session")
async def get_session():
    """Return stored session tokens (if any) so another add-in can restore the login."""
    # Check in-memory cache first (fast path)
    session = _session_store.get("current")
    if session and session.get("access_token"):
        # Update last_active
        session["last_active"] = datetime.utcnow().isoformat()
        _update_last_active_supabase()
        return {"session": session, "source": "memory"}
    # Fallback 1: filesystem backup (survives process restarts within deploy)
    file_data = _load_from_file()
    if file_data and file_data.get("access_token"):
        _session_store["current"] = file_data
        logger.info("[auth] Session restored from filesystem backup")
        return {"session": file_data, "source": "file"}
    # Fallback 2: Supabase (survives Railway redeploys — if table exists)
    sb_data = _load_from_supabase()
    if sb_data and sb_data.get("access_token"):
        # Warm up memory cache
        cached = {
            "access_token": sb_data["access_token"],
            "refresh_token": sb_data["refresh_token"],
            "user_id": sb_data.get("user_id", ""),
            "email": sb_data.get("email", ""),
            "last_active": datetime.utcnow().isoformat(),
        }
        _session_store["current"] = cached
        logger.info("[auth] Session restored from Supabase")
        return {"session": cached, "source": "supabase"}
    logger.info("[auth] No session found in any store")
    return {"session": None}


def _update_last_active_supabase():
    """Update last_active timestamp in Supabase."""
    client = _get_supabase()
    if not client:
        return
    try:
        client.table("sessions").update({"last_active": datetime.utcnow().isoformat()}).eq("id", "current").execute()
    except Exception:
        pass  # Non-critical, don't log


@router.post("/clear-session")
async def clear_session():
    """Clear stored session (on sign-out)."""
    _session_store.pop("current", None)
    _user_store.pop("current", None)
    try:
        SESSION_FILE.unlink(missing_ok=True)
    except Exception:
        pass
    _clear_supabase()
    logger.info("[auth] Session cleared")
    return {"status": "ok"}


# Init-DB endpoint — creates the sessions table
@router.post("/init-db")
async def init_db():
    """Manually trigger sessions table creation. Requires SUPABASE_SERVICE_ROLE_KEY or DATABASE_URL."""
    ok = await ensure_sessions_table()
    if ok:
        return {"status": "ok", "message": "Sessions table created/verified"}
    return {
        "status": "error",
        "message": "Could not create sessions table. Set SUPABASE_SERVICE_ROLE_KEY or DATABASE_URL, or create manually.",
        "sql": (
            "CREATE TABLE IF NOT EXISTS sessions ("
            "id TEXT PRIMARY KEY DEFAULT 'current', "
            "access_token TEXT, refresh_token TEXT, user_id TEXT, email TEXT, "
            "last_active TIMESTAMPTZ DEFAULT now(), updated_at TIMESTAMPTZ DEFAULT now()"
            "); "
            "ALTER TABLE sessions ENABLE ROW LEVEL SECURITY; "
            "CREATE POLICY \"Allow all access\" ON sessions FOR ALL USING (true) WITH CHECK (true);"
        ),
    }


# Auth status endpoint — shows what's working and what's not
@router.get("/status")
async def auth_status():
    """Diagnostic endpoint showing auth system health."""
    has_memory = bool(_session_store.get("current", {}).get("access_token"))
    has_file = bool(_load_from_file())
    has_supabase = False
    supabase_error = None
    try:
        sb_data = _load_from_supabase()
        has_supabase = bool(sb_data and sb_data.get("access_token"))
    except Exception as e:
        supabase_error = str(e)

    return {
        "session_in_memory": has_memory,
        "session_in_file": has_file,
        "session_in_supabase": has_supabase,
        "supabase_table_ok": _supabase_table_ok,
        "supabase_error": supabase_error,
        "service_role_key_set": bool(os.getenv("SUPABASE_SERVICE_ROLE_KEY")),
        "database_url_set": bool(os.getenv("DATABASE_URL")),
        "email": _session_store.get("current", {}).get("email", ""),
    }


# Rate limit login attempts (Improvement 6)
class LoginAttemptRequest(BaseModel):
    email: str

@router.post("/check-login-rate")
async def check_login_rate(req: LoginAttemptRequest):
    """Check if login attempts are rate limited. Returns 429 if too many attempts."""
    email = req.email.lower().strip()
    now = time.time()
    if email in _login_attempts:
        entry = _login_attempts[email]
        # Reset if window expired (60 seconds)
        if now - entry["first_attempt"] > 60:
            _login_attempts[email] = {"count": 1, "first_attempt": now}
            return {"allowed": True}
        entry["count"] += 1
        if entry["count"] > 5:
            raise HTTPException(status_code=429, detail="Too many login attempts. Please wait 60 seconds.")
        return {"allowed": True}
    else:
        _login_attempts[email] = {"count": 1, "first_attempt": now}
        return {"allowed": True}
