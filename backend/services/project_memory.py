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
import re
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
    """Primary (stable) workbook id.

    Prefers `workbook_name` if the client sent it — that's the actual filename
    and survives sheet additions/renames. Falls back to sorted-sheets hash for
    older clients or when the workbook name isn't available (e.g. unsaved
    "Book1" before first save — Office.js returns "" in that case, we fall
    back to the sheets hash).
    """
    ctx = context or {}
    app = ctx.get("app") or ""
    name = (ctx.get("workbook_name") or "").strip()
    if name:
        sig: dict[str, Any] = {"app": app, "workbook": name}
    else:
        sig = {"app": app, "sheets": sorted(ctx.get("all_sheets") or [])}
    raw = json.dumps(sig, sort_keys=True).encode()
    return hashlib.sha256(raw).hexdigest()[:16]


def fingerprint_legacy(context: dict) -> str:
    """Secondary (pre-workbook_name) workbook id. Used as a fallback lookup
    target so state written before the `workbook_name` upgrade isn't orphaned.
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


# ── Supabase adapter (primary) with file fallback ────────────────────────────
# The file path is suitable for local dev but /tmp on Railway is wiped on every
# deploy. Supabase stores the state durably. If SUPABASE_URL/KEY env vars are
# present, use Supabase; otherwise fall back to disk so nothing breaks in dev.

def _supabase_client():
    """Return the shared Supabase client from services.memory, or None."""
    try:
        from services.memory import _get_client
        return _get_client()
    except Exception as e:
        logger.debug("project_memory: supabase client unavailable: %s", e)
        return None


def load(user_id: str, workbook_id: str, state_dir: Path | None = None) -> dict:
    """Load state for (user_id, workbook_id). Returns empty template if none."""
    client = _supabase_client()
    if client is not None:
        try:
            result = client.table("project_memory_state") \
                .select("state") \
                .eq("user_id", user_id) \
                .eq("workbook_id", workbook_id) \
                .execute()
            rows = result.data or []
            if rows:
                return rows[0]["state"] or _empty_state(workbook_id)
            return _empty_state(workbook_id)
        except Exception as e:
            logger.warning("project_memory: supabase load failed, falling back to file: %s", e)

    # File fallback
    path = _state_path(user_id, workbook_id, state_dir)
    if path.exists():
        try:
            return json.loads(path.read_text())
        except json.JSONDecodeError as e:
            logger.warning("project_memory: corrupt state at %s: %s — starting fresh", path, e)
    return _empty_state(workbook_id)


def save(user_id: str, workbook_id: str, state: dict, state_dir: Path | None = None) -> None:
    """Persist state. Supabase first, file fallback on error or if unconfigured."""
    state["updated_at"] = time.time()

    # Bound growth regardless of backend
    if len(state.get("completed", [])) > MAX_COMPLETED_STORED:
        state["completed"] = state["completed"][-MAX_COMPLETED_STORED:]
    if len(state.get("turns", [])) > 50:
        state["turns"] = state["turns"][-50:]

    client = _supabase_client()
    if client is not None:
        try:
            from datetime import datetime, timezone
            client.table("project_memory_state").upsert({
                "user_id":     user_id,
                "workbook_id": workbook_id,
                "state":       state,
                "updated_at":  datetime.now(timezone.utc).isoformat(),
            }, on_conflict="user_id,workbook_id").execute()
            return
        except Exception as e:
            # Loud logging so the fallback doesn't hide real failures.
            logger.error("project_memory: supabase save FAILED, falling back to file: %s: %s",
                         type(e).__name__, e)

    # File fallback
    path = _state_path(user_id, workbook_id, state_dir)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(state, indent=2, default=str))
    tmp.replace(path)


def clear(user_id: str, workbook_id: str, state_dir: Path | None = None) -> bool:
    """Wipe state for a workbook. Returns True if something was deleted."""
    deleted = False

    client = _supabase_client()
    if client is not None:
        try:
            result = client.table("project_memory_state") \
                .delete() \
                .eq("user_id", user_id) \
                .eq("workbook_id", workbook_id) \
                .execute()
            if result.data:
                deleted = True
        except Exception as e:
            logger.warning("project_memory: supabase clear failed, also trying file: %s", e)

    # Also wipe file (harmless if Supabase already cleared it)
    path = _state_path(user_id, workbook_id, state_dir)
    if path.exists():
        path.unlink()
        deleted = True

    return deleted


def backend_type() -> str:
    """For diagnostics: returns 'supabase' or 'file' depending on what's active."""
    return "supabase" if _supabase_client() is not None else "file"


# ── Prompt injection ─────────────────────────────────────────────────────────

_A1_RE = re.compile(r"^\$?([A-Z]+)\$?(\d+)$")


def _col_to_idx(col: str) -> int:
    """A → 0, B → 1, ... AA → 26"""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _cell_is_present_in_context(addr: str, context: dict) -> bool | None:
    """Check whether `Sheet!Cell` is non-empty in the live workbook preview.

    Returns:
      True   — the cell has a value/formula in preview (memory entry still valid)
      False  — the cell is in preview range but empty (memory entry is stale)
      None   — unknown (cell outside preview, sheet missing, no context) —
               caller should treat as "keep" to avoid false invalidation.
    """
    if not addr or "!" not in addr or not context:
        return None
    sheet, cell = addr.split("!", 1)
    sheet = sheet.strip().strip("'")
    cell = cell.strip()

    # Only A1 single-cell addresses for now — ranges are too ambiguous to reconcile
    m = _A1_RE.match(cell)
    if not m:
        return None
    col_idx = _col_to_idx(m.group(1))
    row_idx = int(m.group(2)) - 1

    summaries = context.get("sheet_summaries") or []
    for s in summaries:
        if (s.get("name") or "").casefold() == sheet.casefold():
            preview = s.get("preview") or []
            if row_idx >= len(preview):
                return None  # outside preview rows — unknown
            row = preview[row_idx] or []
            if col_idx >= len(row):
                return None  # outside preview cols — unknown
            val = row[col_idx]
            return val is not None and val != ""
    return None  # sheet not in context


def build_prompt_injection(state: dict, context: dict | None = None) -> str:
    """Format state as a prefix to inject into the user message.

    If `context` is provided, each completed entry is reconciled against the
    live workbook preview — stale entries (cell claimed done but live cell is
    empty) are flagged so the LLM is warned about the drift rather than being
    told to trust stale memory. Empty string if there's nothing useful to say.
    """
    completed = (state.get("completed") or [])[-MAX_COMPLETED_IN_PROMPT:]
    locks     = state.get("user_locks") or []
    turn_count = len(state.get("turns") or [])

    if not completed and not locks:
        return ""

    # Partition completed entries by reconciliation status
    live: list[dict] = []
    stale: list[dict] = []
    unknown: list[dict] = []
    for item in completed:
        addr = item.get("cell") or ""
        status = _cell_is_present_in_context(addr, context or {})
        if status is True:
            live.append(item)
        elif status is False:
            stale.append(item)
        else:
            unknown.append(item)

    lines = [
        "## WORKBOOK STATE FROM PRIOR TURNS",
        "",
        f"You have worked on this same workbook across {turn_count} prior turn(s). "
        "The cells below were already written in prior turns. Use this to avoid "
        "redoing work — focus on what's NOT yet done.",
        "",
    ]

    if live or unknown:
        lines.append("### Already written (leave alone unless asked):")
        for item in live + unknown:
            addr = item.get("cell") or item.get("range") or "?"
            tag  = item.get("type") or ""
            body = item.get("formula") or item.get("note") or item.get("name") or ""
            body = str(body)[:80]
            lines.append(f"- {addr}  ({tag})  {body}")
        lines.append("")

    if stale:
        # Memory claims these cells have content but the live workbook disagrees.
        # Most likely cause: user opened a fresh template or copied the file.
        # Tell the LLM so it doesn't "skip" a cell that actually needs writing.
        lines.append("### Memory drift detected (memory says done, live cell is empty — you may need to RE-EMIT these):")
        for item in stale:
            addr = item.get("cell") or "?"
            body = item.get("formula") or item.get("note") or item.get("name") or ""
            body = str(body)[:80]
            lines.append(f"- {addr}  was {body}")
        lines.append("")

    if locks:
        lines.append("### User-locked ranges (NEVER modify):")
        for item in locks:
            rng  = item.get("range") or "?"
            note = item.get("note") or "user said don't touch"
            lines.append(f"- {rng} — {note}")
        lines.append("")

    lines.append(
        "If you believe a cell needs to change despite being in the 'Already "
        "written' list, explain why in your reply and ask the user to confirm."
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

def load_with_migration(user_id: str, context: dict) -> tuple[str, dict]:
    """Load state under the primary fingerprint; if empty, migrate from the
    legacy (sheets-hash) fingerprint. Returns (workbook_id, state).

    Called by inject_and_load (on every chat turn) AND by the memory-panel
    lookup endpoint, so the panel shows migrated state on boot.
    """
    wb_id = fingerprint(context)
    state = load(user_id, wb_id)

    if not (state.get("completed") or state.get("user_locks")):
        legacy_id = fingerprint_legacy(context)
        if legacy_id != wb_id:
            legacy_state = load(user_id, legacy_id)
            if legacy_state.get("completed") or legacy_state.get("user_locks"):
                logger.info(
                    "project_memory: migrating legacy state %s → %s for user %s",
                    legacy_id, wb_id, user_id,
                )
                legacy_state["workbook_id"] = wb_id
                try:
                    save(user_id, wb_id, legacy_state)
                    clear(user_id, legacy_id)
                    state = legacy_state
                except Exception as e:
                    logger.warning("project_memory: legacy migration failed: %s", e)

    return wb_id, state


def inject_and_load(user_id: str, context: dict, message: str) -> tuple[str, str, dict]:
    """Convenience: fingerprint → load (with migration) → build injection → prepend.

    Returns (new_message, workbook_id, state).
    """
    wb_id, state = load_with_migration(user_id, context)
    if not is_enabled():
        return message, wb_id, state
    injection = build_prompt_injection(state, context=context)
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
