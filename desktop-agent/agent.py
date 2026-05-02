#!/usr/bin/env python3
"""
Tsifl Desktop Agent — local Excel automation controller.

Runs on your Mac, polls the backend for computer use tasks,
and executes them via AppleScript + xlwings (ZERO pyautogui).

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
from pathlib import Path
from dotenv import load_dotenv

# Load API key from project .env
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    load_dotenv(env_path, override=True)

BACKEND_URL = os.getenv("BACKEND_URL", "https://focused-solace-production-6839.up.railway.app")


def _cancel_watcher_loop(http, session_id: str, stop_fn):
    """Background thread: polls backend every 500ms for cancel status."""
    while True:
        try:
            resp = http.get(f"{BACKEND_URL}/computer-use/status/{session_id}")
            if resp.status_code == 200:
                status = resp.json().get("status")
                if status == "cancelled":
                    print(f"[agent] Cancel detected for {session_id}")
                    stop_fn()
                    return
                if status in ("completed", "failed"):
                    return
        except Exception:
            pass
        time.sleep(0.5)


def _is_cancelled(http, session_id: str) -> bool:
    try:
        resp = http.get(f"{BACKEND_URL}/computer-use/status/{session_id}")
        if resp.status_code == 200:
            return resp.json().get("status") == "cancelled"
    except Exception:
        pass
    return False


def poll_and_execute():
    """Main loop: poll backend for pending tasks, execute them."""
    print(f"[agent] Tsifl Desktop Agent starting...")
    print(f"[agent] Backend: {BACKEND_URL}")
    print(f"[agent] Mode: AppleScript + xlwings (zero pyautogui)")

    http = httpx.Client(timeout=30)

    print(f"[agent] Ready. Polling for tasks...")
    print(f"[agent] Press Ctrl+C to stop")
    print()

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

                    print(f"\n[agent] ========================================")
                    print(f"[agent] New task: session {session_id}")
                    print(f"[agent] Actions: {len(actions)}")
                    for a in actions:
                        print(f"[agent]   - {a.get('type')}")
                    print(f"[agent] ========================================\n")

                    # Claim the session
                    http.post(f"{BACKEND_URL}/computer-use/claim/{session_id}")

                    from excel_applescript import (
                        execute_all_actions as applescript_execute,
                        clear_stop, set_stop,
                    )

                    # Start cancel watcher
                    clear_stop()
                    cancel_watcher = threading.Thread(
                        target=_cancel_watcher_loop,
                        args=(http, session_id, set_stop),
                        daemon=True,
                    )
                    cancel_watcher.start()

                    # Execute all actions via AppleScript + xlwings
                    result = applescript_execute(
                        actions, context,
                        cancel_check=lambda: _is_cancelled(http, session_id),
                    )

                    # Stop the cancel watcher
                    set_stop()
                    cancel_watcher.join(timeout=2)
                    clear_stop()

                    if not result:
                        result = {"status": "completed", "message": "No actions", "steps_taken": 0}

                    # If cancelled mid-run, override status
                    if _is_cancelled(http, session_id):
                        result["status"] = "cancelled"
                        result["message"] = "Stopped by user"

                    # Report result
                    http.post(
                        f"{BACKEND_URL}/computer-use/complete/{session_id}",
                        json=result,
                    )

                    print(f"[agent] Session {session_id}: {result['status']}")

        except (httpx.ConnectError, httpx.ReadTimeout, httpx.TimeoutException):
            pass
        except KeyboardInterrupt:
            print("\n[agent] Shutting down...")
            break
        except Exception as e:
            import traceback
            print(f"[agent] Error: {e}")
            traceback.print_exc()

        time.sleep(2)


def main():
    """Entry point with auto-restart on crash."""
    while True:
        try:
            poll_and_execute()
            break  # Clean exit (Ctrl+C)
        except KeyboardInterrupt:
            print("\n[agent] Shutting down...")
            break
        except Exception as e:
            import traceback
            print(f"\n[agent] FATAL: {e}")
            traceback.print_exc()
            print("[agent] Restarting in 3 seconds...")
            time.sleep(3)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--applescript":
        # Direct AppleScript mode: python3 agent.py --applescript '{"type":"goal_seek",...}'
        from excel_applescript import execute_excel_action
        action = json.loads(sys.argv[2])
        print(f"[agent] Direct: {action.get('type')}")
        result = execute_excel_action(action)
        print(json.dumps(result, indent=2))
    else:
        main()
