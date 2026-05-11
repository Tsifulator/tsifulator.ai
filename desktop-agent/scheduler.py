"""
scheduler.py — background daemon that fires due routines.

Wakes every 30s, checks routines.json for any with next_run <= now,
sends their prompt through the v2 agent loop, logs the result, and
schedules the next fire.

Routines run as if the user pressed ⌘⌥T and typed the prompt — they get
the same tools, the same context, the same agent loop. But the panel
stays hidden; the user gets a macOS notification when a routine completes
(unless the result is empty/uninteresting).
"""

from __future__ import annotations
import os
import sys
import time
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional


_TICK_SECONDS = 30
_scheduler_thread: Optional[threading.Thread] = None
_scheduler_stop = threading.Event()


def _now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%H:%M:%S")


def _notify(title: str, message: str):
    """Show a macOS notification — safe to call from a background thread
    when invoked via rumps. We do it best-effort."""
    try:
        import rumps  # type: ignore
        rumps.notification(title=title, subtitle="", message=message[:240])
    except Exception:
        try:
            os.system(
                f'osascript -e \'display notification "{message[:200]}" with title "{title}"\''
            )
        except Exception:
            pass


def _run_one_routine(r):
    """Fire a single routine's prompt through the v2 agent loop."""
    from agent_v2_client import run_agent_loop
    from routines import record_run

    try:
        from executor import get_system_context
        mac_ctx = get_system_context()
    except Exception:
        mac_ctx = {}

    try:
        from memory import get_memory_context
        mem = get_memory_context()
    except Exception:
        mem = ""

    context = {
        "frontmost_app": mac_ctx.get("frontmost_app", ""),
        "mac": mac_ctx,
        "user_memory": mem,
        # Tag so prompts can know this isn't an interactive turn
        "_source": "routine",
        "_routine_name": r.name,
    }

    sys.stderr.write(f"[scheduler {_now_iso()}] firing routine [{r.id}] {r.name!r}\n")

    try:
        summary = run_agent_loop(
            user_message=r.prompt,
            context=context,
            images=None,
            max_steps=5,
        )
    except Exception as e:
        sys.stderr.write(f"[scheduler] {r.id} crashed: {e}\n")
        record_run(r.id, "error", f"crash: {e}")
        _notify(f"tsifl • {r.name}", f"⚠️ failed: {e}")
        return

    if summary.get("error") and summary["error"] not in ("budget_exceeded", "turn_cap_exceeded"):
        record_run(r.id, "error", summary.get("error", "unknown"))
        _notify(f"tsifl • {r.name}", f"⚠️ {summary['error']}")
        return

    final_text = (summary.get("final_text") or "(no output)").strip()
    # Trim very long results for the notification (full result in log)
    notif_body = final_text[:200] + ("…" if len(final_text) > 200 else "")
    record_run(r.id, "ok", final_text)
    _notify(f"tsifl • {r.name}", notif_body)


def _tick():
    """One pass — find due routines and run them. Safe to call repeatedly."""
    from routines import due_routines, record_run

    try:
        due = due_routines()
    except Exception as e:
        sys.stderr.write(f"[scheduler] due_routines crashed: {e}\n")
        return

    if not due:
        return

    sys.stderr.write(f"[scheduler {_now_iso()}] {len(due)} routine(s) due\n")
    for r in due:
        # Mark "next_run" forward FIRST so we don't double-fire if the
        # run is slow. record_run() rolls next_run on completion too,
        # but doing it upfront protects against overlap.
        try:
            record_run(r.id, r.last_status or "running", "scheduled — running")
        except Exception:
            pass
        _run_one_routine(r)


def _scheduler_loop():
    sys.stderr.write("[scheduler] started\n")
    while not _scheduler_stop.is_set():
        try:
            _tick()
        except Exception as e:
            sys.stderr.write(f"[scheduler] tick crashed: {e}\n")
        # Sleep but wake early if stop requested
        _scheduler_stop.wait(_TICK_SECONDS)
    sys.stderr.write("[scheduler] stopped\n")


def start_scheduler() -> bool:
    """Start the daemon thread. Idempotent — safe to call multiple times."""
    global _scheduler_thread
    if _scheduler_thread is not None and _scheduler_thread.is_alive():
        return True
    _scheduler_stop.clear()
    _scheduler_thread = threading.Thread(
        target=_scheduler_loop, name="tsifl-scheduler", daemon=True,
    )
    _scheduler_thread.start()
    return True


def stop_scheduler():
    """Signal the loop to exit. Daemon will join on process exit anyway."""
    _scheduler_stop.set()


def run_routine_now(routine_id_or_name: str) -> tuple[bool, str]:
    """Trigger a routine immediately (without waiting for the schedule)."""
    from routines import find_routine

    r = find_routine(routine_id_or_name)
    if not r:
        return False, f"No routine matched '{routine_id_or_name}'."
    # Run in a thread so we don't block the caller
    threading.Thread(target=_run_one_routine, args=(r,), daemon=True).start()
    return True, f"Triggered routine [{r.id}] {r.name!r}"
