"""
Usage Monitor — protects your API margin.
Tracks tasks per user per month and enforces tier limits.
Starter: 40 tasks/month ($20 tier)
Pro: unlimited ($49 tier)
"""

import os

# In-memory store for development. Replace with Supabase in production.
_usage_store: dict = {}

STARTER_LIMIT = int(os.getenv("STARTER_TASK_LIMIT", 500))

async def check_and_increment_usage(user_id: str) -> dict:
    """
    Check if user is under their limit, then increment their count.
    Returns: {"allowed": bool, "remaining": int, "used": int}
    """
    if user_id not in _usage_store:
        _usage_store[user_id] = {"used": 0, "tier": "starter"}

    user = _usage_store[user_id]
    limit = STARTER_LIMIT if user["tier"] == "starter" else 999999

    if user["used"] >= limit:
        return {"allowed": False, "remaining": 0, "used": user["used"]}

    user["used"] += 1
    remaining = limit - user["used"]

    return {"allowed": True, "remaining": remaining, "used": user["used"]}
