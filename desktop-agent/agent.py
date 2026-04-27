#!/usr/bin/env python3
"""
Tsifl Desktop Agent — the local screen controller.
Runs on your Mac, polls the backend for computer use tasks,
and executes them by controlling your screen via Claude's Computer Use API.

Usage:
  python3 agent.py

Requirements:
  pip3 install anthropic pyautogui Pillow httpx

Architecture:
  1. Polls backend /computer-use/pending for tasks
  2. Takes a screenshot of your screen
  3. Sends screenshot + instructions to Claude (computer-use model)
  4. Claude says "click here" / "type this"
  5. Agent executes the action on your screen
  6. Takes new screenshot, sends back to Claude
  7. Repeats until Claude says "done"
  8. Reports completion to backend
"""

import anthropic
import pyautogui
import httpx
import base64
import time
import json
import os
import sys
import subprocess
import tempfile
import threading
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from dotenv import load_dotenv

# Load API key from project .env. Walk up the tree from this file's location
# until we find a .env or hit the home dir. This means the agent works whether
# it's run from the main repo (~/tsifulator.ai/.env) or from a worktree
# (~/tsifulator.ai/.claude/worktrees/X/.env) — without manual symlinks.
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

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
BACKEND_URL = os.getenv("BACKEND_URL", "https://focused-solace-production-6839.up.railway.app")
CU_MODEL = "claude-sonnet-4-20250514"

# Safety: disable pyautogui failsafe (move mouse to corner to abort)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3  # small delay between actions

# ── Persistent logging ──────────────────────────────────────────────────────
# Writes to both stdout (preserves existing "[agent] ..." UX) AND a rotating
# log file at ~/Library/Logs/tsifl-agent.log so we have post-hoc debug trails
# when the agent is bundled as a .app and stdout disappears.
LOG_DIR = Path.home() / "Library" / "Logs"
try:
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    LOG_FILE = LOG_DIR / "tsifl-agent.log"
except Exception:
    # Fallback if home dir isn't writable for some weird reason
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
        # Don't crash the agent if the file logger can't open — console is enough
        print(f"[agent] WARN: could not open {LOG_FILE}: {_log_err}", file=sys.stderr)


def _log(msg: str, level: str = "info") -> None:
    """Single entry point for agent logging. Use this instead of print()."""
    getattr(logger, level, logger.info)(msg)

# Screen dimensions (will be detected)
SCREEN_W = 0
SCREEN_H = 0


def take_screenshot() -> str:
    """Take a screenshot, resize to fit API limits, return as base64 JPEG."""
    from PIL import Image
    import io

    tmp = tempfile.mktemp(suffix=".png")
    try:
        # macOS screencapture
        subprocess.run(
            ["screencapture", "-x", "-C", tmp],
            check=True,
            capture_output=True,
        )
        img = Image.open(tmp)
    except Exception as e:
        _log(f"screencapture failed: {e}, trying pyautogui...")
        img = pyautogui.screenshot()
    finally:
        try:
            os.unlink(tmp)
        except OSError:
            pass

    # Convert RGBA to RGB (JPEG doesn't support alpha)
    if img.mode == "RGBA":
        img = img.convert("RGB")

    # Use native resolution (Retina macs capture at 2x, so cap at 1920)
    max_width = 1920
    if img.width > max_width:
        ratio = max_width / img.width
        new_size = (max_width, int(img.height * ratio))
        img = img.resize(new_size, Image.LANCZOS)

    # Save as JPEG (much smaller than PNG)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=70)
    data = buf.getvalue()
    _log(f"Screenshot: {img.width}x{img.height}, {len(data)//1024}KB")
    return base64.standard_b64encode(data).decode("utf-8")


def get_screen_size():
    """Get the screen dimensions."""
    global SCREEN_W, SCREEN_H
    try:
        size = pyautogui.size()
        SCREEN_W = size.width
        SCREEN_H = size.height
    except Exception:
        SCREEN_W = 1920
        SCREEN_H = 1080
    _log(f"Screen size: {SCREEN_W}x{SCREEN_H}")


def _scale_coord(coord, img_w, img_h):
    """Scale coordinates from image space to logical screen space.
    Retina Macs: image might be larger than logical screen coordinates."""
    x, y = coord[0], coord[1]
    scale_x = SCREEN_W / img_w if img_w > 0 else 1
    scale_y = SCREEN_H / img_h if img_h > 0 else 1
    return int(x * scale_x), int(y * scale_y)

# Will be set per-session to the screenshot dimensions
_img_w = 1920
_img_h = 1080


def execute_computer_action(action: str, input_data: dict) -> str:
    """Execute a single computer use action on the local screen.
    Returns a status message."""
    try:
        if action == "screenshot":
            return "screenshot_taken"

        elif action == "mouse_move":
            coord = input_data.get("coordinate", [0, 0])
            sx, sy = _scale_coord(coord, _img_w, _img_h)
            pyautogui.moveTo(sx, sy, duration=0.2)
            return f"Mouse moved to ({coord[0]}, {coord[1]})"

        elif action == "left_click":
            coord = input_data.get("coordinate", [0, 0])
            sx, sy = _scale_coord(coord, _img_w, _img_h)
            pyautogui.click(sx, sy)
            return f"Clicked at ({coord[0]}, {coord[1]}) → screen ({sx}, {sy})"

        elif action == "right_click":
            coord = input_data.get("coordinate", [0, 0])
            sx, sy = _scale_coord(coord, _img_w, _img_h)
            pyautogui.rightClick(sx, sy)
            return f"Right-clicked at ({coord[0]}, {coord[1]})"

        elif action == "double_click":
            coord = input_data.get("coordinate", [0, 0])
            sx, sy = _scale_coord(coord, _img_w, _img_h)
            pyautogui.doubleClick(sx, sy)
            return f"Double-clicked at ({coord[0]}, {coord[1]})"

        elif action == "left_click_drag":
            start = input_data.get("start_coordinate", [0, 0])
            end = input_data.get("coordinate", [0, 0])
            s1x, s1y = _scale_coord(start, _img_w, _img_h)
            s2x, s2y = _scale_coord(end, _img_w, _img_h)
            pyautogui.moveTo(s1x, s1y, duration=0.1)
            pyautogui.mouseDown()
            pyautogui.moveTo(s2x, s2y, duration=0.3)
            pyautogui.mouseUp()
            return f"Dragged from ({start[0]},{start[1]}) to ({end[0]},{end[1]})"

        elif action == "type":
            text = input_data.get("text", "")
            pyautogui.write(text, interval=0.02)
            return f"Typed: {text[:50]}"

        elif action == "key":
            key = input_data.get("text", "")
            # Handle modifier combos like "ctrl+a", "cmd+s"
            key = key.replace("Return", "enter").replace("return", "enter")
            key = key.replace("Tab", "tab").replace("Escape", "escape")
            if "+" in key:
                parts = key.split("+")
                # Convert to pyautogui hotkey
                mapped = []
                for p in parts:
                    p = p.strip().lower()
                    p = p.replace("ctrl", "command" if sys.platform == "darwin" else "ctrl")
                    p = p.replace("super", "command" if sys.platform == "darwin" else "win")
                    mapped.append(p)
                pyautogui.hotkey(*mapped)
            else:
                pyautogui.press(key.lower())
            return f"Pressed: {key}"

        elif action == "scroll":
            coord = input_data.get("coordinate", [SCREEN_W // 2, SCREEN_H // 2])
            direction = input_data.get("direction", "down")
            amount = input_data.get("amount", 3)
            pyautogui.moveTo(coord[0], coord[1])
            if direction == "up":
                pyautogui.scroll(amount)
            elif direction == "down":
                pyautogui.scroll(-amount)
            elif direction == "left":
                pyautogui.hscroll(-amount)
            elif direction == "right":
                pyautogui.hscroll(amount)
            return f"Scrolled {direction} by {amount} at ({coord[0]},{coord[1]})"

        elif action == "wait":
            time.sleep(input_data.get("duration", 1))
            return "Waited"

        else:
            _log(f"Unknown action: {action}")
            return f"Unknown action: {action}"

    except Exception as e:
        _log(f"Action '{action}' failed: {e}")
        return f"Action failed: {e}"


# ── CU loop tuning constants ────────────────────────────────────────────────
# Lowered from 50 → 12. Stuck tasks now fail in ~2 minutes instead of 10+,
# saving ~$6 per stuck task. The vast majority of legitimate Excel ribbon
# tasks finish in 4-8 iterations; 12 leaves headroom without burning money.
CU_MAX_ITERATIONS = 12

# Wall-clock timeout per CU task. The model can take 5-15 seconds per
# iteration, so 90s ≈ 6-12 iterations. Combined with CU_MAX_ITERATIONS this
# gives a hard upper bound regardless of which limit fires first.
CU_TASK_TIMEOUT_SECONDS = 90

# When Claude clicks the same coordinate (within ±10px) this many times in a
# row with no resolution, we abort. This is the smoking gun for "model is
# confused and looping" — it almost never recovers from that state.
CU_REPEATED_CLICK_LIMIT = 3

# System prompt sent ALONGSIDE the user's instructions. Anthropic's CU model
# accepts a system prompt; previously we sent none, so the model defaulted to
# "click anything that looks plausible". This prompt enforces fail-fast.
_CU_SYSTEM_PROMPT = """You are tsifl's desktop helper, controlling a Mac with Microsoft Excel for an analyst.

OPERATING RULES — these supersede general computer-use intuition:

1. COORDINATES ARE EXACT PIXEL VALUES. Click precisely on the pixel I tell you. Do not guess "near" a button — read the exact pixel position from the screenshot.

2. NO REPEATED CLICKING. If you click an element and the screen does not visibly change after the next screenshot, do NOT click the same element again. The click landed but had no effect — either the button is disabled, you misidentified the element, or the dialog is in a different state than you expected. Press Esc or Cmd+Z first, then take a different approach.

3. THREE STRIKES RULE. After 3 consecutive failed attempts at the same goal (clicking same area, same dialog, same field), STOP and report failure. Do not loop trying the same thing 30 times — it costs the user money and never works.

4. COMMIT TO ONE PATH. Do not jump between menu paths mid-task. If you opened the Data menu, complete the operation there. Don't open Insert → close → reopen Data → close → try Format. Pick a route, finish it, or report failure.

5. VERIFY BEFORE TYPING. Before typing into a field, take a screenshot and confirm the cursor is in the correct field. Typing into the wrong field is unrecoverable without Cmd+Z.

6. SHORT TASKS COMPLETE IN UNDER 8 STEPS. Most Excel ribbon operations (Solver setup, Data Table, Goal Seek) require 4-8 actions: open dialog, fill fields, click OK. If you find yourself on iteration 9+ for a single operation, you are stuck — report failure instead of pressing on.

7. NEVER CLOSE OR INTERACT WITH NON-EXCEL WINDOWS. The user has other apps open. Stay focused on Excel.

8. ON FAILURE, BE HONEST. If the dialog isn't there, the menu item doesn't exist in this Excel version, or the screen state is wrong, say so plainly: "I cannot complete this — the [X] dialog did not open after [Y]." Do not pretend success.

When done, stop emitting tool calls and write a brief one-sentence summary of what you accomplished or why you stopped.
"""


def _coords_match(a, b, tol: int = 10) -> bool:
    """True if two [x, y] coordinates are within `tol` pixels of each other."""
    if not isinstance(a, (list, tuple)) or not isinstance(b, (list, tuple)):
        return False
    if len(a) < 2 or len(b) < 2:
        return False
    return abs(a[0] - b[0]) <= tol and abs(a[1] - b[1]) <= tol


def _safe_back_out() -> None:
    """Try to leave Excel in a clean state when we abort.

    Press Esc twice (closes any open dialog/popover). Best-effort — never
    raise. Called on stop / timeout / repeated-click abort so the user
    doesn't end up with a stuck modal dialog blocking their workbook.
    """
    try:
        pyautogui.press("escape")
        time.sleep(0.1)
        pyautogui.press("escape")
    except Exception as e:
        _log(f"safe_back_out: pyautogui Esc failed: {e}", "warning")


def run_computer_use_loop(
    instructions: str,
    session_id: str,
    stop_check=None,
):
    """Run the full Claude computer-use loop for a task.

    Args:
        instructions: Task description for the model.
        session_id: For logging only.
        stop_check: Optional callable, called BEFORE every API call. If it
            returns True, the loop aborts gracefully (Esc back-out + return
            "cancelled" status). Caller is responsible for plumbing this in.

    Hard limits enforced:
        - CU_MAX_ITERATIONS (12): caps step count
        - CU_TASK_TIMEOUT_SECONDS (90): caps wall-clock time
        - CU_REPEATED_CLICK_LIMIT (3): aborts on stuck-click pattern

    Returns dict with `status` ∈ {completed, failed, cancelled, timeout, stuck},
    `message`, `steps_taken`.
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    start_time = time.time()
    stop_check = stop_check or (lambda: False)

    # Initial screenshot
    _log(f"[{session_id}] Taking initial screenshot...")
    screenshot_b64 = take_screenshot()

    messages = [
        {
            "role": "user",
            "content": [
                {"type": "text", "text": instructions},
                {
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": screenshot_b64,
                    },
                },
            ],
        }
    ]

    # Use the SCREENSHOT dimensions for display size (what Claude sees)
    # not pyautogui's logical screen size
    from PIL import Image
    import io as _io
    _tmp_b64 = base64.standard_b64decode(screenshot_b64)
    _tmp_img = Image.open(_io.BytesIO(_tmp_b64))
    display_w, display_h = _tmp_img.size
    _log(f"[{session_id}] Display for Claude: {display_w}x{display_h}")

    # Set globals for coordinate scaling
    global _img_w, _img_h
    _img_w = display_w
    _img_h = display_h

    tools = [
        {
            "type": "computer_20250124",
            "name": "computer",
            "display_width_px": display_w,
            "display_height_px": display_h,
            "display_number": 1,
        }
    ]

    iteration = 0

    # Recent clickable-action history for stuck-loop detection.
    # Stores tuples of (action_type, coord) for click-like actions only.
    recent_clicks: list[tuple[str, list]] = []

    while iteration < CU_MAX_ITERATIONS:
        # Stop check #1: before the (expensive) API call
        if stop_check():
            _log(f"[{session_id}] Stop signal detected before API call (iter {iteration})")
            _safe_back_out()
            return {
                "status": "cancelled",
                "message": "Stopped by user",
                "steps_taken": iteration,
            }

        # Hard wall-clock timeout
        elapsed = time.time() - start_time
        if elapsed >= CU_TASK_TIMEOUT_SECONDS:
            _log(f"[{session_id}] Timeout after {elapsed:.1f}s (iter {iteration})", "warning")
            _safe_back_out()
            return {
                "status": "timeout",
                "message": f"Timed out after {int(elapsed)}s without completing",
                "steps_taken": iteration,
            }

        iteration += 1
        _log(f"[{session_id}] === Iteration {iteration}/{CU_MAX_ITERATIONS} (t+{elapsed:.1f}s) ===")

        try:
            response = client.beta.messages.create(
                model=CU_MODEL,
                max_tokens=4096,
                system=_CU_SYSTEM_PROMPT,
                tools=tools,
                messages=messages,
                betas=["computer-use-2025-01-24"],
            )
        except Exception as e:
            _log(f"[{session_id}] API error: {e}", "error")
            _safe_back_out()
            return {"status": "failed", "error": str(e), "steps_taken": iteration}

        # Process response
        tool_uses = [b for b in response.content if b.type == "tool_use"]
        text_blocks = [b for b in response.content if b.type == "text"]

        for tb in text_blocks:
            _log(f"[{session_id}] Claude: {tb.text[:200]}")

        if not tool_uses:
            # Claude is done
            final_text = " ".join(b.text for b in text_blocks)
            _log(f"[{session_id}] Task completed in {iteration} iterations")
            return {
                "status": "completed",
                "message": final_text,
                "steps_taken": iteration,
            }

        # Execute each tool use and collect results
        tool_results = []
        for tu in tool_uses:
            # Stop check #2: between tool executions inside an iteration.
            # If the user hits Stop while Claude returned 5 tool_uses in
            # one response, we abort after the current one rather than
            # plowing through all 5.
            if stop_check():
                _log(f"[{session_id}] Stop signal mid-iteration", "warning")
                _safe_back_out()
                return {
                    "status": "cancelled",
                    "message": "Stopped by user",
                    "steps_taken": iteration,
                }

            action = tu.input.get("action", "")
            _log(f"[{session_id}] Action: {action} | {json.dumps(tu.input)[:150]}")

            # Stuck-loop detection: track click-like actions and bail if
            # the model is hitting the same coordinate repeatedly.
            if action in ("left_click", "right_click", "double_click"):
                coord = tu.input.get("coordinate", [0, 0])
                recent_clicks.append((action, coord))
                # Keep only the last N for matching window
                if len(recent_clicks) > CU_REPEATED_CLICK_LIMIT:
                    recent_clicks = recent_clicks[-CU_REPEATED_CLICK_LIMIT:]
                if (len(recent_clicks) == CU_REPEATED_CLICK_LIMIT
                    and all(rc[0] == action for rc in recent_clicks)
                    and all(_coords_match(rc[1], coord) for rc in recent_clicks)):
                    _log(
                        f"[{session_id}] STUCK: {CU_REPEATED_CLICK_LIMIT} "
                        f"identical {action}s at ~{coord} — aborting",
                        "warning",
                    )
                    _safe_back_out()
                    return {
                        "status": "stuck",
                        "message": (
                            f"Model clicked the same spot ({coord}) "
                            f"{CU_REPEATED_CLICK_LIMIT}× without progress. "
                            "Aborted to avoid infinite loop."
                        ),
                        "steps_taken": iteration,
                    }
            else:
                # Non-click action breaks the streak
                recent_clicks = []

            # Execute the action (wrapped — pyautogui can raise on permission denies)
            try:
                result_text = execute_computer_action(action, tu.input)
            except Exception as e:
                _log(f"[{session_id}] execute_computer_action raised: {e}", "error")
                result_text = f"Action raised: {e}"

            # Take a fresh screenshot after the action
            time.sleep(0.5)  # brief pause for screen to update
            try:
                new_screenshot = take_screenshot()
            except Exception as e:
                _log(f"[{session_id}] post-action screenshot failed: {e}", "error")
                # Last-ditch: skip the screenshot, send only text result
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": tu.id,
                    "content": [{"type": "text", "text": result_text}],
                })
                continue

            # Build tool result with screenshot
            tool_result_content = []
            if action == "screenshot" or result_text == "screenshot_taken":
                tool_result_content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": new_screenshot,
                    },
                })
            else:
                tool_result_content.append({
                    "type": "text",
                    "text": result_text,
                })
                tool_result_content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": new_screenshot,
                    },
                })

            tool_results.append({
                "type": "tool_result",
                "tool_use_id": tu.id,
                "content": tool_result_content,
            })

        # Feed back to Claude
        messages.append({"role": "assistant", "content": response.content})
        messages.append({"role": "user", "content": tool_results})

    # Loop fell out without Claude saying "done" — hit iteration cap
    _log(f"[{session_id}] Hit max iterations ({CU_MAX_ITERATIONS}) without completing", "warning")
    _safe_back_out()
    return {
        "status": "failed",
        "error": f"Max iterations ({CU_MAX_ITERATIONS}) reached without completion",
        "steps_taken": iteration,
    }


def _cancel_watcher_loop(http, session_id: str, stop_fn):
    """Background thread: polls backend every 500ms for cancel status.
    When detected, calls stop_fn() to set the global stop flag."""
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
                    return  # Already done, no need to watch
        except Exception:
            pass
        time.sleep(0.5)


def _is_cancelled(http, session_id: str) -> bool:
    """Check if the backend session has been cancelled by the user."""
    try:
        resp = http.get(f"{BACKEND_URL}/computer-use/status/{session_id}")
        if resp.status_code == 200:
            return resp.json().get("status") == "cancelled"
    except Exception:
        pass
    return False


def poll_and_execute():
    """Main loop: poll backend for pending tasks, execute them."""
    _log(f"Tsifl Desktop Agent starting...")
    _log(f"Backend: {BACKEND_URL}")
    _log(f"Model: {CU_MODEL}")
    get_screen_size()

    if not ANTHROPIC_API_KEY:
        _log("ERROR: ANTHROPIC_API_KEY not set!")
        sys.exit(1)

    http = httpx.Client(timeout=10)

    _log(f"Ready. Polling for computer use tasks...")
    _log(f"Press Ctrl+C to stop")
    print()

    while True:
        try:
            # Check for pending sessions
            resp = http.get(f"{BACKEND_URL}/computer-use/pending")
            if resp.status_code == 200:
                data = resp.json()
                sessions = data.get("sessions", [])

                for session in sessions:
                    session_id = session["id"]
                    actions = session.get("actions", [])
                    context = session.get("context", {})

                    print(f"\n[agent] ========================================")
                    _log(f"New task: session {session_id}")
                    _log(f"Actions: {len(actions)}")
                    _log(f"========================================\n")

                    # Claim the session
                    http.post(
                        f"{BACKEND_URL}/computer-use/claim/{session_id}"
                    )

                    # PRIMARY: Use AppleScript for known Excel operations
                    # FALLBACK: Use computer use only for unknown actions
                    from excel_applescript import (
                        execute_all_actions as applescript_execute,
                        clear_stop, set_stop,
                    )

                    known_types = {
                        "create_data_table", "goal_seek", "scenario_manager",
                        "scenario_summary", "run_solver", "save_solver_scenario",
                        "run_toolpak", "install_addins", "uninstall_addins",
                    }

                    known_actions = [a for a in actions if a.get("type") in known_types]
                    unknown_actions = [a for a in actions if a.get("type") not in known_types]

                    # Start background cancel watcher — polls every 500ms
                    # Sets the global stop flag so check_stop() raises mid-action
                    clear_stop()
                    cancel_watcher = threading.Thread(
                        target=_cancel_watcher_loop,
                        args=(http, session_id, set_stop),
                        daemon=True,
                    )
                    cancel_watcher.start()

                    result = None

                    if known_actions:
                        _log(f"AppleScript path: {len(known_actions)} known actions")
                        result = applescript_execute(
                            known_actions, context,
                            cancel_check=lambda: _is_cancelled(http, session_id),
                        )

                    if unknown_actions and not _is_cancelled(http, session_id):
                        _log(f"Computer Use fallback: {len(unknown_actions)} unknown actions")
                        from services_local import build_instructions
                        instructions = build_instructions(unknown_actions, context)
                        # Pass stop_check so the CU loop checks for cancellation
                        # BEFORE every API call, not just between iterations
                        cu_result = run_computer_use_loop(
                            instructions, session_id,
                            stop_check=lambda: _is_cancelled(http, session_id),
                        )
                        if result:
                            result["message"] += f" | CU: {cu_result.get('message', '')}"
                            if cu_result["status"] == "failed":
                                result["status"] = "partial"
                        else:
                            result = cu_result

                    # Stop the cancel watcher
                    set_stop()
                    cancel_watcher.join(timeout=2)
                    clear_stop()

                    if not result:
                        result = {"status": "completed", "message": "No actions to execute", "steps_taken": 0}

                    # If cancelled mid-run, override status
                    if _is_cancelled(http, session_id):
                        result["status"] = "cancelled"
                        result["message"] = "Stopped by user"
                        _log(f"Session {session_id} was cancelled")

                    # Report result back to backend
                    http.post(
                        f"{BACKEND_URL}/computer-use/complete/{session_id}",
                        json=result,
                    )

                    _log(f"Session {session_id}: {result['status']}")

        except httpx.ConnectError:
            pass  # Backend not reachable, retry
        except KeyboardInterrupt:
            print("\n[agent] Shutting down...")
            break
        except Exception as e:
            _log(f"Error: {e}")

        time.sleep(2)  # Poll every 2 seconds


# ── Standalone mode ──────────────────────────────────────────────────────────

def run_standalone(task_description: str):
    """Run a single computer use task directly (no backend needed)."""
    _log(f"Standalone mode: {task_description}")
    get_screen_size()

    if not ANTHROPIC_API_KEY:
        _log("ERROR: ANTHROPIC_API_KEY not set!")
        sys.exit(1)

    instructions = (
        "You are controlling a Mac with Microsoft Excel open.\n\n"
        "RULES:\n"
        "- First click on the Excel window in the taskbar/dock to bring it to focus\n"
        "- Act FAST — do NOT take extra screenshots unless needed to verify\n"
        "- Click precisely on buttons and menu items\n"
        "- Do NOT try to close or interact with other windows\n"
        "- On macOS Excel, the ribbon tabs are: Home, Insert, Draw, Page Layout, Formulas, Data, Review, View, Automate\n"
        "- 'What-If Analysis' is in the Data tab, in the 'Forecast' section on the right side\n\n"
        f"Task: {task_description}\n\n"
        "Start by taking a screenshot."
    )

    result = run_computer_use_loop(instructions, "standalone")
    print(f"\n[agent] Result: {json.dumps(result, indent=2)}")
    return result


def run_applescript_standalone(action_json: str):
    """Run a structured action via AppleScript (no backend, no computer use).
    Pass a JSON string like: '{"type":"create_data_table","payload":{"range":"B14:E23","col_input_cell":"$B$4","sheet":"Calorie Journal"}}'
    """
    from excel_applescript import execute_excel_action
    import json as _json

    action = _json.loads(action_json)
    _log(f"AppleScript standalone: {action.get('type')}")
    result = execute_excel_action(action)
    print(f"\n[agent] Result: {_json.dumps(result, indent=2)}")
    return result


if __name__ == "__main__":
    if len(sys.argv) > 1:
        first_arg = sys.argv[1]

        if first_arg == "--applescript" and len(sys.argv) > 2:
            # AppleScript mode: python3 agent.py --applescript '{"type":"create_data_table",...}'
            run_applescript_standalone(sys.argv[2])
        elif first_arg == "--test-menu":
            # Quick test: python3 agent.py --test-menu
            from excel_applescript import activate_excel, open_data_table_dialog
            print("[test] Activating Excel...")
            activate_excel()
            time.sleep(1)
            print("[test] Opening Data Table dialog via menu bar...")
            open_data_table_dialog()
            print("[test] Done — check if the Data Table dialog opened in Excel")
        elif first_arg == "--test-all":
            # Test all operations: python3 agent.py --test-all
            from excel_applescript import activate_excel, search_menu
            print("[test] Testing Help menu search...")
            activate_excel()
            time.sleep(0.5)
            for term in ["Goal Seek", "Scenario Manager", "Solver", "Data Analysis"]:
                print(f"[test] Searching for: {term}")
                search_menu(term)
                time.sleep(1)
                import pyautogui
                pyautogui.press('escape')
                time.sleep(0.5)
                pyautogui.press('escape')
                time.sleep(0.5)
            print("[test] All menu searches completed")
        else:
            # Computer Use mode: python3 agent.py "Create a data table"
            task = " ".join(sys.argv[1:])
            run_standalone(task)
    else:
        poll_and_execute()
