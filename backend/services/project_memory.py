"""Per-project memory — remembers what's already correct across turns.

The problem this solves:
  Every tsifl run today is stateless. If turn 1 correctly writes F5 and
  turn 2 gets the same prompt, the LLM re-runs the whole project from
  scratch and may write F5 wrong this time. Progress doesn't accumulate.

The fix:
  After each turn, record which cells were written with which formulas.
  On the next turn for the same workbook, inject this into the prompt
  so the LLM knows "these cells already have the right content — don't
  touch them unless explicitly asked."

Scope:
  File-based per-(user_id, workbook_id) state. Suitable for MVP. Should
  migrate to Supabase or a proper store when we outgrow local JSON.

Enable/disable via env var PROJECT_MEMORY_ENABLED=1. Default off so this
ships behind a flag and doesn't affect existing behavior until we're
ready.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import time
import uuid
from pathlib import Path
from typing import Any

logger = logging.getLogger("project_memory")

# Where state lives. Railway may or may not have persistent volumes — for MVP
# we accept ephemerality; state rebuilds itself as the user interacts. If we
# outgrow this, migrate to Supabase.
DEFAULT_STATE_DIR = Path(os.environ.get("TSIFL_STATE_DIR", "/tmp/tsifl-state"))

# How many recent completed items to inject into a prompt. Too many = token bloat.
MAX_COMPLETED_IN_PROMPT = 30

# Trim completed list to this size on write — keeps disk usage bounded.
MAX_COMPLETED_STORED = 200


def is_enabled() -> bool:
    """Feature flag. Disabled by default."""
    return os.environ.get("PROJECT_MEMORY_ENABLED", "").strip() in ("1", "true", "yes", "on")


# ── Workbook fingerprinting ──────────────────────────────────────────────────

def fingerprint(context: dict) -> str:
    """Stable short id for a workbook across turns.

    Based on app + sorted sheet names. This is intentionally coarse — it
    groups turns against the SAME workbook structure together. If the user
    renames sheets or switches to a different workbook, the fingerprint
    changes and state is independent.

    A more precise fingerprint (e.g. including first-N-cells-per-sheet) would
    guard against false matches but risk state loss on small edits. Start
    simple; tighten if collisions show up.
    """
    sig: dict[str, Any] = {
        "app":    (context or {}).get("app") or "",
        "sheets": sorted((context or {}).get("all_sheets") or []),
    }
    raw = json.dumps(sig, sort_keys=True).encode()
    return hashlib.sha256(raw).hexdigest()[:16]


# ── Storage ──────────────────────────────────────────────────────────────────

def _safe_id(s: str) -> str:
    """Make an id filesystem-safe."""
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in (s or ""))[:128] or "anon"


def _state_path(user_id: str, workbook_id: str, state_dir: Path | None = None) -> Path:
    root = state_dir or DEFAULT_STATE_DIR
    return root / _safe_id(user_id) / f"{_safe_id(workbook_id)}.json"


def _empty_state(workbook_id: str) -> dict:
    return {
        "workbook_id": workbook_id,
        "created_at":  time.time(),
        "updated_at":  time.time(),
        "completed":   [],    # [{cell, formula|note, type, turn_id, at}]
        "user_locks":  [],    # [{range, note, at}]
        "turns":       [],    # [{turn_id, at, action_count, reply_preview}]
    }


def load(user_id: str, workbook_id: str, state_dir: Path | None = None) -> dict:
    """Load state for (user_id, workbook_id). Returns empty template if none."""
    path = _state_path(user_id, workbook_id, state_dir)
    if path.exists():
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError as e:
            logger.warning("project_memory: corrupt state at %s: %s — starting fresh", path, e)
    return _empty_state(workbook_id)


def save(user_id: str, workbook_id: str, state: dict, state_dir: Path | None = None) -> None:
    """Persist state atomically."""
    path = _state_path(user_id, workbook_id, state_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    state["updated_at"] = time.time()

    # Bound growth
    if len(state.get("completed", [])) > MAX_COMPLETED_STORED:
        state["completed"] = state["completed"][-MAX_COMPLETED_STORED:]
    if len(state.get("turns", [])) > 50:
        state["turns"] = state["turns"][-50:]

    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2, default=str))
    tmp.replace(path)


def clear(user_id: str, workbook_id: str, state_dir: Path | None = None) -> bool:
    """Wipe state for a workbook. Returns True if something was deleted."""
    path = _state_path(user_id, workbook_id, state_dir)
    if path.exists():
        path.unlink()
        return True
    return False


# ── Prompt injection ─────────────────────────────────────────────────────────

def build_prompt_injection(state: dict) -> str:
    """Format state as a prefix to inject into the user message.

    Empty string if there's nothing useful to say. Keeps it tight — the
    LLM context window is precious.
    """
    completed = (state.get("completed") or [])[-MAX_COMPLETED_IN_PROMPT:]
    locks     = state.get("user_locks") or []
    turn_count = len(state.get("turns") or [])

    if not completed and not locks:
        return ""

    lines = [
        "## WORKBOOK STATE FROM PRIOR TURNS",
        "",
        f"You have worked on this same workbook across {turn_count} prior turn(s). "
        "The cells below already hold correct content. Do NOT overwrite them "
        "unless the user explicitly asks. Focus on what's NOT yet done.",
        "",
    ]

    if completed:
        lines.append("### Already written (leave these alone):")
        for item in completed:
            addr = item.get("cell") or item.get("range") or "?"
            tag  = item.get("type") or ""
            body = item.get("formula") or item.get("note") or item.get("name") or ""
            body = str(body)[:80]
            lines.append(f"- {addr}  ({tag})  {body}")
        lines.append("")

    if locks:
        lines.append("### User-locked ranges (NEVER modify):")
        for item in locks:
            rng  = item.get("range") or "?"
            note = item.get("note") or "user said don't touch"
            lines.append(f"- {rng} — {note}")
        lines.append("")

    lines.append(
        "If you believe a cell needs to change despite being in the list above, "
        "explain why in your reply and ask the user to confirm before acting."
    )
    lines.append("")
    return "\n".join(lines)


# ── Recording actions after a turn ───────────────────────────────────────────

def record_actions(state: dict, actions: list[dict], reply: str = "") -> dict:
    """Append emitted actions to `completed`. Dedupes by target address so
    repeated writes to the same cell keep only the latest formula."""
    turn_id = uuid.uuid4().hex[:8]
    now     = time.time()

    new_completed = []
    for a in actions or []:
        t = a.get("type", "")
        p = a.get("payload") or {}
        s = p.get("sheet")

        if t in ("write_cell", "write_formula"):
            cell = p.get("cell") or p.get("address")
            if not cell:
                continue
            new_completed.append({
                "cell":    f"{s}!{cell}" if s else cell,
                "formula": p.get("formula") or p.get("value"),
                "type":    t,
                "turn_id": turn_id,
                "at":      now,
            })
        elif t == "write_range":
            rng = p.get("range")
            if not rng:
                continue
            new_completed.append({
                "range":   f"{s}!{rng}" if s else rng,
                "note":    "range written",
                "type":    t,
                "turn_id": turn_id,
                "at":      now,
            })
        elif t == "create_named_range":
            new_completed.append({
                "cell":    p.get("reference") or p.get("range"),
                "name":    p.get("name"),
                "note":    f"named range: {p.get('name')}",
                "type":    t,
                "turn_id": turn_id,
                "at":      now,
            })
        elif t == "add_chart":
            anchor = p.get("anchor") or p.get("position")
            new_completed.append({
                "range":   f"{s}!{anchor}" if (s and anchor) else (anchor or "chart"),
                "note":    f"chart: {p.get('title', '')}"[:80],
                "type":    t,
                "turn_id": turn_id,
                "at":      now,
            })
        elif t == "format_range":
            rng = p.get("range")
            if not rng:
                continue
            new_completed.append({
                "range":   f"{s}!{rng}" if s else rng,
                "note":    "formatted",
                "type":    t,
                "turn_id": turn_id,
                "at":      now,
            })

    # Merge with dedup (keep latest per address)
    existing = state.get("completed") or []
    seen: dict[str, dict] = {}
    for item in existing + new_completed:
        addr = item.get("cell") or item.get("range") or item.get("note")
        if addr:
            seen[addr] = item
    state["completed"] = list(seen.values())

    # Log the turn
    turns = state.get("turns") or []
    turns.append({
        "turn_id":        turn_id,
        "at":             now,
        "action_count":   len(actions or []),
        "reply_preview":  (reply or "")[:120],
    })
    state["turns"] = turns

    return state


# ── Lock management (for future "lock F5" user commands) ─────────────────────

def add_lock(state: dict, rng: str, note: str = "") -> dict:
    locks = state.get("user_locks") or []
    if not any(l.get("range") == rng for l in locks):
        locks.append({"range": rng, "note": note, "at": time.time()})
    state["user_locks"] = locks
    return state


def remove_lock(state: dict, rng: str) -> dict:
    state["user_locks"] = [l for l in (state.get("user_locks") or []) if l.get("range") != rng]
    return state


# ── Orchestration helpers (used by chat.py) ──────────────────────────────────

def inject_and_load(user_id: str, context: dict, message: str) -> tuple[str, str, dict]:
    """Convenience: fingerprint → load → build injection → prepend to message.

    Returns (new_message, workbook_id, state). If memory disabled or no state,
    returns (message, workbook_id, empty_state) so caller can still record.
    """
    wb_id = fingerprint(context)
    state = load(user_id, wb_id)
    if not is_enabled():
        return message, wb_id, state
    injection = build_prompt_injection(state)
    if not injection:
        return message, wb_id, state
    return f"{injection}\n---\n\n{message}", wb_id, state


def record_and_save(user_id: str, workbook_id: str, state: dict,
                    actions: list[dict], reply: str = "") -> None:
    """Convenience: record actions + save. Silently skips if disabled."""
    if not is_enabled():
        return
    state = record_actions(state, actions, reply)
    save(user_id, workbook_id, state)
