"""
Auth Route — syncs session across all tsifl add-ins (Excel, Word, PowerPoint, etc.)
Each Office add-in has isolated localStorage, so Supabase sessions don't share.
This module stores the session centrally so any add-in can restore it.

Uses in-memory storage (primary) + filesystem backup.
In-memory survives within a single container lifecycle.
Filesystem backup survives restarts on local dev.
"""

import json
from fastapi import APIRouter
from pydantic import BaseModel
from pathlib import Path

router = APIRouter()

# In-memory session store — primary storage that survives across requests
_session_store = {}
_user_store = {}

# Filesystem backup paths (for local dev persistence)
SESSION_PATH = Path.home() / ".tsifulator_session"
USER_PATH = Path.home() / ".tsifulator_user"


class UserConfig(BaseModel):
    user_id: str
    email: str = ""


class SessionConfig(BaseModel):
    access_token: str
    refresh_token: str
    user_id: str = ""
    email: str = ""


def _save_to_file(path, data):
    """Best-effort filesystem backup."""
    try:
        path.write_text(json.dumps(data) if isinstance(data, dict) else str(data))
    except Exception:
        pass


def _load_from_file(path):
    """Best-effort filesystem restore."""
    try:
        if path.exists():
            text = path.read_text().strip()
            if text.startswith("{"):
                return json.loads(text)
            return text
    except Exception:
        pass
    return None


@router.post("/set-user")
async def set_user(config: UserConfig):
    """Store user ID centrally."""
    _user_store["current"] = config.user_id
    _save_to_file(USER_PATH, config.user_id)
    return {"status": "ok", "user_id": config.user_id}


@router.get("/current-user")
async def current_user():
    """Read the currently saved user ID."""
    uid = _user_store.get("current")
    if not uid:
        uid = _load_from_file(USER_PATH)
        if uid and isinstance(uid, str):
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
    # Store in memory (primary)
    _session_store["current"] = data
    # Backup to file
    _save_to_file(SESSION_PATH, data)
    # Keep user file in sync
    if config.user_id:
        _user_store["current"] = config.user_id
        _save_to_file(USER_PATH, config.user_id)
    return {"status": "ok"}


@router.get("/get-session")
async def get_session():
    """Return stored session tokens (if any) so another add-in can restore the login."""
    # Check in-memory first
    session = _session_store.get("current")
    if session and session.get("access_token"):
        return {"session": session}
    # Fallback to filesystem
    file_data = _load_from_file(SESSION_PATH)
    if file_data and isinstance(file_data, dict) and file_data.get("access_token"):
        _session_store["current"] = file_data  # Warm up memory cache
        return {"session": file_data}
    return {"session": None}


@router.post("/clear-session")
async def clear_session():
    """Clear stored session (on sign-out)."""
    _session_store.pop("current", None)
    _user_store.pop("current", None)
    try:
        if SESSION_PATH.exists():
            SESSION_PATH.unlink()
        if USER_PATH.exists():
            USER_PATH.unlink()
    except Exception:
        pass
    return {"status": "ok"}
