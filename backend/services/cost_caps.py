"""
cost_caps.py — per-user daily spend tracking + caps for the v2 agent.

In-memory; cheap; resets at UTC midnight. Designed for the "burn $25/day"
runaway-loop scenario the user was worried about — bound the blast radius
before it bites.

For production multi-server deployments, swap the in-memory dict for a
Redis backend (interface is the same).
"""

from __future__ import annotations
import os
import threading
from datetime import datetime, timezone


# Soft & hard caps (USD). Override via env so different tiers can have
# different limits without a deploy.
SOFT_CAP_USD = float(os.environ.get("TSIFL_SOFT_CAP_USD", "1.0"))
HARD_CAP_USD = float(os.environ.get("TSIFL_HARD_CAP_USD", "3.0"))

# Per-turn cap: max cost a SINGLE user message can spend across all its
# follow-up rounds (search → read → write → reply, etc.). This prevents
# one runaway prompt from burning the daily budget in one shot.
PER_TURN_CAP_USD = float(os.environ.get("TSIFL_PER_TURN_CAP_USD", "0.25"))


_lock = threading.Lock()
# {user_id: {"date": "YYYY-MM-DD", "spent": float, "calls": int}}
_usage: dict[str, dict] = {}


def _today() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d")


def _reset_if_new_day(entry: dict):
    today = _today()
    if entry.get("date") != today:
        entry["date"] = today
        entry["spent"] = 0.0
        entry["calls"] = 0


def get_daily_spend(user_id: str) -> dict:
    """Returns {"date", "spent", "calls", "remaining_soft", "remaining_hard"}."""
    with _lock:
        entry = _usage.setdefault(user_id, {"date": _today(), "spent": 0.0, "calls": 0})
        _reset_if_new_day(entry)
        return {
            "date": entry["date"],
            "spent": round(entry["spent"], 6),
            "calls": entry["calls"],
            "remaining_soft": max(0.0, SOFT_CAP_USD - entry["spent"]),
            "remaining_hard": max(0.0, HARD_CAP_USD - entry["spent"]),
        }


def check_budget(user_id: str) -> tuple[bool, str]:
    """Returns (allowed, message). Called BEFORE making the API request.

    allowed=False means hard-cap exceeded — refuse the request entirely.
    """
    with _lock:
        entry = _usage.setdefault(user_id, {"date": _today(), "spent": 0.0, "calls": 0})
        _reset_if_new_day(entry)
        spent = entry["spent"]
    if spent >= HARD_CAP_USD:
        return False, (
            f"💰 Daily spend limit reached ({spent:.2f} of {HARD_CAP_USD:.2f} USD). "
            f"Resets at UTC midnight."
        )
    return True, ""


def record_spend(user_id: str, cost_usd: float) -> dict:
    """Called AFTER each API request; returns the post-call status.

    Soft-cap warning surfaces in the response; hard-cap blocks the next request.
    """
    with _lock:
        entry = _usage.setdefault(user_id, {"date": _today(), "spent": 0.0, "calls": 0})
        _reset_if_new_day(entry)
        entry["spent"] += max(0.0, cost_usd)
        entry["calls"] += 1
        spent = entry["spent"]
        calls = entry["calls"]

    warning = ""
    if spent >= HARD_CAP_USD:
        warning = (
            f"💰 You've hit today's hard limit (${spent:.2f}). "
            f"Further requests will be blocked until UTC midnight."
        )
    elif spent >= SOFT_CAP_USD:
        warning = (
            f"⚠️ Approaching daily cap: ${spent:.2f} of ${HARD_CAP_USD:.2f} used."
        )

    return {
        "spent": round(spent, 6),
        "calls": calls,
        "soft_cap": SOFT_CAP_USD,
        "hard_cap": HARD_CAP_USD,
        "warning": warning,
    }
