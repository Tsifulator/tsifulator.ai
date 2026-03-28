"""
Memory Service — Persistent cross-session memory via Supabase.
Stores conversation history and model context so Claude remembers
what was built in previous sessions.
"""

import os
import json
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path(__file__).parent.parent.parent / ".env", override=True)

# Import Supabase client — only connects when credentials are set
_client = None

def _get_client():
    global _client
    if _client is not None:
        return _client

    url = os.getenv("SUPABASE_URL", "")
    key = os.getenv("SUPABASE_KEY", "")

    if not url or not key or url == "your_supabase_project_url":
        return None  # Memory not yet configured — runs fine without it

    from supabase import create_client
    _client = create_client(url, key)
    return _client


async def save_message(user_id: str, role: str, content: str,
                       app: str = "excel", session_id: str = ""):
    """Save a message to persistent memory."""
    client = _get_client()
    if not client:
        return  # Silently skip if Supabase not configured

    try:
        client.table("messages").insert({
            "user_id":    user_id,
            "role":       role,
            "content":    content,
            "app":        app,
            "session_id": session_id,
            "created_at": datetime.utcnow().isoformat()
        }).execute()
    except Exception:
        pass  # Never crash the main flow over memory


async def get_recent_history(user_id: str, limit: int = 10) -> list:
    """
    Fetch recent conversation history for a user across ALL apps.
    This is the shared memory — Excel and RStudio both see the same history.
    """
    client = _get_client()
    if not client:
        return []

    try:
        result = client.table("messages") \
            .select("role, content, app, created_at") \
            .eq("user_id", user_id) \
            .order("created_at", desc=True) \
            .limit(limit) \
            .execute()

        # Return in chronological order for Claude
        return list(reversed(result.data or []))
    except Exception:
        return []


async def save_model_context(user_id: str, model_type: str, context: dict):
    """
    Save a named model snapshot (e.g. 'lbo_model', 'dcf_model').
    Claude can recall this in future sessions.
    """
    client = _get_client()
    if not client:
        return

    try:
        # Upsert — overwrite if same user + model_type exists
        client.table("model_contexts").upsert({
            "user_id":    user_id,
            "model_type": model_type,
            "context":    json.dumps(context),
            "updated_at": datetime.utcnow().isoformat()
        }).execute()
    except Exception:
        pass


async def get_model_context(user_id: str, model_type: str) -> dict:
    """Retrieve a previously saved model context."""
    client = _get_client()
    if not client:
        return {}

    try:
        result = client.table("model_contexts") \
            .select("context") \
            .eq("user_id", user_id) \
            .eq("model_type", model_type) \
            .limit(1) \
            .execute()

        if result.data:
            return json.loads(result.data[0]["context"])
    except Exception:
        pass

    return {}


def is_connected() -> bool:
    """Returns True if Supabase is configured and reachable."""
    return _get_client() is not None
