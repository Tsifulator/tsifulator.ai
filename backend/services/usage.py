"""
Usage Monitor — protects your API margin.
Tracks tasks per user per month and enforces tier limits.
Starter: 40 tasks/month ($20 tier)
Pro: unlimited ($49 tier)
"""

import os
from datetime import datetime

# In-memory store (cache). Supabase provides persistence.
_usage_store: dict = {}

STARTER_LIMIT = int(os.getenv("STARTER_TASK_LIMIT", 500))

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
    except Exception:
        return None


def _get_month_key():
    return datetime.utcnow().strftime("%Y-%m")


async def check_and_increment_usage(user_id: str) -> dict:
    """
    Check if user is under their limit, then increment their count.
    Returns: {"allowed": bool, "remaining": int, "used": int}
    """
    month_key = _get_month_key()
    cache_key = f"{user_id}:{month_key}"

    # Try Supabase first for persistence
    client = _get_supabase()
    if client:
        try:
            result = client.table("usage").select("*").eq("user_id", user_id).eq("month", month_key).execute()
            if result.data:
                used = result.data[0].get("used", 0)
                tier = result.data[0].get("tier", "starter")
            else:
                used = 0
                tier = "starter"

            limit = STARTER_LIMIT if tier == "starter" else 999999
            if used >= limit:
                return {"allowed": False, "remaining": 0, "used": used}

            new_used = used + 1
            client.table("usage").upsert({
                "user_id": user_id,
                "month": month_key,
                "used": new_used,
                "tier": tier,
            }).execute()
            return {"allowed": True, "remaining": limit - new_used, "used": new_used}
        except Exception:
            pass  # Fall through to in-memory

    # In-memory fallback
    if cache_key not in _usage_store:
        _usage_store[cache_key] = {"used": 0, "tier": "starter"}

    user = _usage_store[cache_key]
    limit = STARTER_LIMIT if user["tier"] == "starter" else 999999

    if user["used"] >= limit:
        return {"allowed": False, "remaining": 0, "used": user["used"]}

    user["used"] += 1
    remaining = limit - user["used"]

    return {"allowed": True, "remaining": remaining, "used": user["used"]}
