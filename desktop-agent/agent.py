#!/usr/bin/env python3
"""
Tsifl Desktop Agent — local Excel automation controller.

Pure AppleScript + xlwings. ZERO pyautogui. ZERO screen takeover.

Runs in the background while the user keeps working in other apps.
Polls the backend for tasks, dispatches them to excel_applescript handlers
which talk to Excel directly via the macOS scripting bridge or xlwings COM.

Usage:
  python3 agent.py

Requirements:
  pip3 install xlwings httpx python-dotenv
"""

import httpx
import time
import json
import os
import sys
import threading
import logging
import tempfile
from logging.handlers import RotatingFileHandler
from pathlib import Path
from dotenv import load_dotenv


# ── Env discovery (walk up the tree until we find .env) ─────────────────────
# Works whether the agent is run from the main repo, a worktree, or a
# bundled .app — finds the project root's .env automatically.
def _find_env_file() -> Path | None:
    here = Path(__file__).resolve().parent
    home = Path.home().resolve()
    for candidate in [here, *here.parents]:
        env = candidate / ".env"
        if env.exists():
            return env
        if candidate == home or candidate == candidate.parent:
            break
    return None


env_path = _find_env_file()
if env_path:
    load_dotenv(env_path, override=True)


BACKEND_URL = os.getenv("BACKEND_URL", "https://focused-solace-production-6839.up.railway.app")


# ── Persistent logging ──────────────────────────────────────────────────────
# Both stdout (for terminal-run UX) AND a rotating file at
# ~/Library/Logs/tsifl-agent.log so the .app bundle (where stdout vanishes)
# still has post-hoc debug trails.
LOG_DIR = Path.home() / "Library" / "Logs"
try:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE = LOG_DIR / "tsifl-agent.log"
except Exception:
    LOG_FILE = Path(tempfile.gettempdir()) / "tsifl-agent.log"

logger = logging.getLogger("tsifl-agent")
logger.setLevel(logging.INFO)
logger.propagate = False
if not logger.handlers:  # idempotent — survive double-import
    _console = logging.StreamHandler(sys.stdout)
    _console.setFormatter(logging.Formatter("[agent] %(message)s"))
    logger.addHandler(_console)
    try:
        _file = RotatingFileHandler(
            LOG_FILE, maxBytes=2_000_000, backupCount=3, encoding="utf-8"
        )
        _file.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(_file)
    except Exception as _log_err:
        print(f"[agent] WARN: could not open {LOG_FILE}: {_log_err}", file=sys.stderr)


def _log(msg: str, level: str = "info") -> None:
    """Single entry point for agent logging. Use this instead of print()."""
    getattr(logger, level, logger.info)(msg)


# ── Cancel-watcher: backgrounds polling /computer-use/status for 'cancelled' ─
# When detected, sets the global stop flag in excel_applescript so any
# in-flight AppleScript can raise StopAutomation at its next check_stop().

def _cancel_watcher_loop(http, session_id: str, stop_fn):
    """Polls every 500ms; calls stop_fn() when status flips to 'cancelled'."""
    while True:
        try:
            resp = http.get(f"{BACKEND_URL}/computer-use/status/{session_id}")
            if resp.status_code == 200:
                status = resp.json().get("status")
                if status == "cancelled":
                    _log(f"Cancel detected for {session_id} — setting stop flag")
                    stop_fn()
                    return
                if status in ("completed", "failed"):
                    return  # task already done, no need to watch
        except Exception:
            pass
        time.sleep(0.5)


def _is_cancelled(http, session_id: str) -> bool:
    """Quick blocking check — used inside execute_all_actions between actions."""
    try:
        resp = http.get(f"{BACKEND_URL}/computer-use/status/{session_id}")
        if resp.status_code == 200:
            return resp.json().get("status") == "cancelled"
    except Exception:
        pass
    return False


# ── Main poll loop ──────────────────────────────────────────────────────────

def poll_and_execute():
    """Main loop: poll backend for pending tasks, execute via AppleScript+xlwings."""
    _log("Tsifl Desktop Agent starting...")
    _log(f"Backend: {BACKEND_URL}")
    _log("Mode: AppleScript + xlwings (zero pyautogui, runs in background)")
    _log(f"Logs: {LOG_FILE}")

    http = httpx.Client(timeout=30)

    _log("Ready. Polling for tasks...")
    _log("Press Ctrl+C to stop")

    while True:
        try:
            resp = http.get(f"{BACKEND_URL}/computer-use/pending")
            if resp.status_code == 200:
                data = resp.json()
                sessions = data.get("sessions", [])

                for session in sessions:
                    session_id = session["id"]
                    actions = session.get("actions", [])
                    context = session.get("context", {})

                    _log("=" * 60)
                    _log(f"New task: session {session_id}, {len(actions)} action(s)")
                    for a in actions:
                        _log(f"  - {a.get('type')}")
                    _log("=" * 60)

                    # Claim the session so other agents (if any) skip it
                    http.post(f"{BACKEND_URL}/computer-use/claim/{session_id}")

                    # Lazy import so a missing xlwings install doesn't kill
                    # the agent on startup — only fails when a task actually
                    # needs it. clear_stop / set_stop manage the cancel flag
                    # that excel_applescript checks between AppleScript ops.
                    from excel_applescript import (
                        execute_all_actions as applescript_execute,
                        clear_stop, set_stop,
                    )

                    # Background cancel watcher — polls /computer-use/status
                    # every 500ms and sets the stop flag if user clicks Stop
                    clear_stop()
                    cancel_watcher = threading.Thread(
                        target=_cancel_watcher_loop,
                        args=(http, session_id, set_stop),
                        daemon=True,
                    )
                    cancel_watcher.start()

                    # Dispatch all actions to the AppleScript+xlwings layer.
                    # Unknown action types return {"status": "unknown", ...}
                    # which we surface to the user instead of silently failing.
                    result = applescript_execute(
                        actions, context,
                        cancel_check=lambda: _is_cancelled(http, session_id),
                    )

                    # Stop the cancel watcher cleanly
                    set_stop()
                    cancel_watcher.join(timeout=2)
                    clear_stop()

                    if not result:
                        result = {
                            "status": "completed",
                            "message": "No actions to execute",
                            "steps_taken": 0,
                        }

                    # If user cancelled mid-run, override status so backend +
                    # frontend both see the cancellation cleanly
                    if _is_cancelled(http, session_id):
                        result["status"] = "cancelled"
                        result["message"] = "Stopped by user"
                        _log(f"Session {session_id} was cancelled")

                    # Report result back to backend → frontend access modal
                    # closes and the user sees what happened
                    http.post(
                        f"{BACKEND_URL}/computer-use/complete/{session_id}",
                        json=result,
                    )

                    _log(f"Session {session_id}: {result.get('status', 'unknown')}")

        except (httpx.ConnectError, httpx.ReadTimeout, httpx.TimeoutException):
            # Transient network — backend asleep, slow, etc. Just keep polling.
            pass
        except KeyboardInterrupt:
            _log("Shutting down...")
            break
        except Exception as e:
            import traceback
            _log(f"poll_and_execute error: {e}", "error")
            _log(traceback.format_exc(), "error")

        time.sleep(2)


def main():
    """Entry point with auto-restart on crash. Survives transient errors so
    the .app bundle doesn't need user intervention if something goes wrong."""
    while True:
        try:
            poll_and_execute()
            break  # Clean exit (Ctrl+C)
        except KeyboardInterrupt:
            _log("Shutting down...")
            break
        except Exception as e:
            import traceback
            _log(f"FATAL: {e}", "error")
            _log(traceback.format_exc(), "error")
            _log("Restarting in 3 seconds...")
            time.sleep(3)


# ── CLI shortcuts for dev/debugging ──────────────────────────────────────────

def run_applescript_standalone(action_json: str):
    """Run a single structured action directly via AppleScript+xlwings.
    Useful for: python3 agent.py --applescript '{"type":"goal_seek",...}'
    """
    from excel_applescript import execute_excel_action
    action = json.loads(action_json)
    _log(f"Direct: {action.get('type')}")
    result = execute_excel_action(action)
    print(json.dumps(result, indent=2))
    return result


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--applescript" and len(sys.argv) > 2:
        run_applescript_standalone(sys.argv[2])
    else:
        main()
