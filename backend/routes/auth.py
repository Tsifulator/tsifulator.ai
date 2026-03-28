"""
Auth Route — syncs the logged-in user ID to a local config file.
This lets the RStudio addin pick up the same user identity as Excel.
"""

from fastapi import APIRouter
from pydantic import BaseModel
from pathlib import Path

router = APIRouter()

class UserConfig(BaseModel):
    user_id: str
    email: str = ""

@router.post("/set-user")
async def set_user(config: UserConfig):
    """Write user ID to ~/.tsifulator_user so RStudio addin can read it."""
    config_path = Path.home() / ".tsifulator_user"
    config_path.write_text(config.user_id)
    return {"status": "ok", "user_id": config.user_id}

@router.get("/current-user")
async def current_user():
    """Read the currently saved user ID."""
    config_path = Path.home() / ".tsifulator_user"
    if config_path.exists():
        return {"user_id": config_path.read_text().strip()}
    return {"user_id": None}
