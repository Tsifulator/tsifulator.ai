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
from pathlib import Path
from dotenv import load_dotenv

# Load API key from project .env
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    load_dotenv(env_path, override=True)

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
BACKEND_URL = os.getenv("BACKEND_URL", "https://focused-solace-production-6839.up.railway.app")
CU_MODEL = "claude-sonnet-4-20250514"

# Safety: disable pyautogui failsafe (move mouse to corner to abort)
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3  # small delay between actions

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
        print(f"[agent] screencapture failed: {e}, trying pyautogui...")
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
    print(f"[agent] Screenshot: {img.width}x{img.height}, {len(data)//1024}KB")
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
    print(f"[agent] Screen size: {SCREEN_W}x{SCREEN_H}")


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
            print(f"[agent] Unknown action: {action}")
            return f"Unknown action: {action}"

    except Exception as e:
        print(f"[agent] Action '{action}' failed: {e}")
        return f"Action failed: {e}"


def run_computer_use_loop(instructions: str, session_id: str):
    """Run the full Claude computer-use loop for a task.

    1. Take screenshot
    2. Send to Claude with instructions + computer tool
    3. Claude responds with tool_use (click/type/etc.)
    4. Execute action
    5. Take new screenshot
    6. Repeat until Claude stops using tools
    """
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # Initial screenshot
    print(f"[agent] Taking initial screenshot...")
    screenshot_b64 = take_screenshot()

    messages = [
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": instructions,
                },
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
    print(f"[agent] Display for Claude: {display_w}x{display_h}")

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

    max_iterations = 50
    iteration = 0

    while iteration < max_iterations:
        iteration += 1
        print(f"[agent] === Iteration {iteration} ===")

        try:
            response = client.beta.messages.create(
                model=CU_MODEL,
                max_tokens=4096,
                tools=tools,
                messages=messages,
                betas=["computer-use-2025-01-24"],
            )
        except Exception as e:
            print(f"[agent] API error: {e}")
            return {"status": "failed", "error": str(e), "steps_taken": iteration}

        # Process response
        tool_uses = [b for b in response.content if b.type == "tool_use"]
        text_blocks = [b for b in response.content if b.type == "text"]

        for tb in text_blocks:
            print(f"[agent] Claude says: {tb.text[:200]}")

        if not tool_uses:
            # Claude is done
            final_text = " ".join(b.text for b in text_blocks)
            print(f"[agent] Task completed in {iteration} iterations")
            return {
                "status": "completed",
                "message": final_text,
                "steps_taken": iteration,
            }

        # Execute each tool use and collect results
        tool_results = []
        for tu in tool_uses:
            action = tu.input.get("action", "")
            print(f"[agent] Action: {action} | {json.dumps(tu.input)[:150]}")

            # Execute the action
            result_text = execute_computer_action(action, tu.input)

            # Take a fresh screenshot after the action
            time.sleep(0.5)  # brief pause for screen to update
            new_screenshot = take_screenshot()

            # Build tool result with screenshot
            tool_result_content = []
            if action == "screenshot" or result_text == "screenshot_taken":
                # Just the screenshot
                tool_result_content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": new_screenshot,
                    },
                })
            else:
                # Action result + new screenshot
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

    return {"status": "failed", "error": "Max iterations reached", "steps_taken": iteration}


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
    print(f"[agent] Tsifl Desktop Agent starting...")
    print(f"[agent] Backend: {BACKEND_URL}")
    print(f"[agent] Model: {CU_MODEL}")
    get_screen_size()

    if not ANTHROPIC_API_KEY:
        print("[agent] ERROR: ANTHROPIC_API_KEY not set!")
        sys.exit(1)

    http = httpx.Client(timeout=10)

    print(f"[agent] Ready. Polling for computer use tasks...")
    print(f"[agent] Press Ctrl+C to stop")
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
                    print(f"[agent] New task: session {session_id}")
                    print(f"[agent] Actions: {len(actions)}")
                    print(f"[agent] ========================================\n")

                    # Claim the session
                    http.post(
                        f"{BACKEND_URL}/computer-use/claim/{session_id}"
                    )

                    # PRIMARY: Use AppleScript for known Excel operations
                    # FALLBACK: Use computer use only for unknown actions
                    from excel_applescript import execute_all_actions as applescript_execute

                    known_types = {
                        "create_data_table", "goal_seek", "scenario_manager",
                        "scenario_summary", "run_solver", "save_solver_scenario",
                        "run_toolpak", "install_addins", "uninstall_addins",
                    }

                    known_actions = [a for a in actions if a.get("type") in known_types]
                    unknown_actions = [a for a in actions if a.get("type") not in known_types]

                    result = None

                    if known_actions:
                        print(f"[agent] AppleScript path: {len(known_actions)} known actions")
                        result = applescript_execute(
                            known_actions, context,
                            cancel_check=lambda: _is_cancelled(http, session_id),
                        )

                    if unknown_actions and not _is_cancelled(http, session_id):
                        print(f"[agent] Computer Use fallback: {len(unknown_actions)} unknown actions")
                        from services_local import build_instructions
                        instructions = build_instructions(unknown_actions, context)
                        cu_result = run_computer_use_loop(instructions, session_id)
                        if result:
                            result["message"] += f" | CU: {cu_result.get('message', '')}"
                            if cu_result["status"] == "failed":
                                result["status"] = "partial"
                        else:
                            result = cu_result

                    if not result:
                        result = {"status": "completed", "message": "No actions to execute", "steps_taken": 0}

                    # If cancelled mid-run, override status
                    if _is_cancelled(http, session_id):
                        result["status"] = "cancelled"
                        result["message"] = "Stopped by user"
                        print(f"[agent] Session {session_id} was cancelled")

                    # Report result back to backend
                    http.post(
                        f"{BACKEND_URL}/computer-use/complete/{session_id}",
                        json=result,
                    )

                    print(f"[agent] Session {session_id}: {result['status']}")

        except httpx.ConnectError:
            pass  # Backend not reachable, retry
        except KeyboardInterrupt:
            print("\n[agent] Shutting down...")
            break
        except Exception as e:
            print(f"[agent] Error: {e}")

        time.sleep(2)  # Poll every 2 seconds


# ── Standalone mode ──────────────────────────────────────────────────────────

def run_standalone(task_description: str):
    """Run a single computer use task directly (no backend needed)."""
    print(f"[agent] Standalone mode: {task_description}")
    get_screen_size()

    if not ANTHROPIC_API_KEY:
        print("[agent] ERROR: ANTHROPIC_API_KEY not set!")
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
    print(f"[agent] AppleScript standalone: {action.get('type')}")
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
