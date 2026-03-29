"""
Auth Route — syncs session across all tsifl add-ins (Excel, Word, PowerPoint, etc.)
Each Office add-in has isolated localStorage, so Supabase sessions don't share.
This module stores the session centrally so any add-in can restore it.
"""

import json
from fastapi import APIRouter
from pydantic import BaseModel
from pathlib import Path

router = APIRouter()

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


@router.post("/set-user")
async def set_user(config: UserConfig):
    """Write user ID to ~/.tsifulator_user so RStudio addin can read it."""
    USER_PATH.write_text(config.user_id)
    return {"status": "ok", "user_id": config.user_id}


@router.get("/current-user")
async def current_user():
    """Read the currently saved user ID."""
    if USER_PATH.exists():
        return {"user_id": USER_PATH.read_text().strip()}
    return {"user_id": None}


@router.post("/set-session")
async def set_session(config: SessionConfig):
    """Store Supabase session tokens so all add-ins can share the login."""
    data = {
        "access_token": config.access_token,
        "refresh_token": config.refresh_token,
        "user_id": config.user_id,
        "email": config.email,
    }
    SESSION_PATH.write_text(json.dumps(data))
    # Also keep user file in sync
    if config.user_id:
        USER_PATH.write_text(config.user_id)
    return {"status": "ok"}


@router.get("/get-session")
async def get_session():
    """Return stored session tokens (if any) so another add-in can restore the login."""
    if SESSION_PATH.exists():
        try:
            data = json.loads(SESSION_PATH.read_text())
            return {"session": data}
        except (json.JSONDecodeError, KeyError):
            pass
    return {"session": None}


@router.post("/clear-session")
async def clear_session():
    """Clear stored session (on sign-out)."""
    if SESSION_PATH.exists():
        SESSION_PATH.unlink()
    if USER_PATH.exists():
        USER_PATH.unlink()
    return {"status": "ok"}
