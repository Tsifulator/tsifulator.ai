"""
Computer Use Service — the desktop automation brain.
Uses Anthropic's Claude computer-use to handle Excel features
that Office.js cannot access: Scenario Manager, Solver, Goal Seek,
Data Tables, Analysis ToolPak.

Architecture:
  1. Backend receives an action that requires GUI interaction
  2. This service spawns a computer-use session with Claude
  3. Claude sees the screen, clicks menus, fills dialogs
  4. When done, it reports completion + any extracted results
  5. The add-in polls for status and picks up results
"""

import anthropic
import os
import json
import uuid
import asyncio
import time
import logging
from typing import Optional
from dotenv import load_dotenv
from pathlib import Path

logger = logging.getLogger(__name__)

load_dotenv(Path(__file__).parent.parent.parent / ".env", override=True)

# Computer Use model — must use Claude with computer-use capability
CU_MODEL = "claude-sonnet-4-20250514"

# Track active sessions
_sessions: dict = {}  # session_id -> session state


# ── Action Classification ──────────────────────────────────────────────────
# Actions that MUST go through Computer Use (Office.js can't do these)
COMPUTER_USE_ACTIONS = {
    "scenario_manager",      # Create/edit scenarios
    "save_solver_scenario",  # Save Solver results as scenario
    "run_solver",            # Run Solver
    "goal_seek",             # Goal Seek
    "create_data_table",     # What-If Data Table (the wizard, not formulas)
    "scenario_summary",      # Generate scenario summary report
    "run_toolpak",           # Analysis ToolPak (descriptive stats, etc.)
    "install_addins",        # Install Excel add-ins (Solver, ToolPak)
    "uninstall_addins",      # Uninstall Excel add-ins
    "computer_use",          # Generic computer use fallback
}

# Actions handled by the Office.js add-in (fast path)
ADDIN_ACTIONS = {
    "navigate_sheet", "write_cell", "write_formula", "write_range",
    "fill_down", "fill_right", "copy_range",
    "create_named_range", "sort_range", "set_number_format",
    "autofit", "autofit_columns", "format_range",
    "add_sheet", "clear_range", "freeze_panes",
    "add_chart", "add_data_validation", "add_conditional_format",
    "import_csv", "save_workbook",
}


def classify_action(action: dict) -> str:
    """Classify an action as 'addin' or 'computer_use'."""
    action_type = action.get("type", "")
    if action_type in COMPUTER_USE_ACTIONS:
        return "computer_use"
    return "addin"


def split_actions(actions: list) -> tuple[list, list]:
    """Split actions into add-in actions and computer-use actions.
    Returns (addin_actions, computer_use_actions)."""
    addin = []
    cu = []
    for action in actions:
        if classify_action(action) == "computer_use":
            cu.append(action)
        else:
            addin.append(action)
    return addin, cu


# ── Computer Use Session Management ────────────────────────────────────────

def create_session(actions: list, context: dict) -> str:
    """Create a new computer use session for the given actions.
    Returns a session_id that can be polled for status."""
    session_id = str(uuid.uuid4())[:8]
    _sessions[session_id] = {
        "id": session_id,
        "status": "pending",       # pending -> running -> completed / failed
        "actions": actions,
        "context": context,
        "created_at": time.time(),
        "result": None,
        "error": None,
        "steps": [],               # log of computer use steps taken
    }
    print(f"[computer_use] Created session {session_id} with {len(actions)} actions")
    return session_id


def get_session(session_id: str) -> Optional[dict]:
    """Get the current state of a computer use session."""
    return _sessions.get(session_id)


async def execute_session(session_id: str):
    """Execute a computer use session asynchronously.
    This is the main loop that drives Claude's computer use."""
    session = _sessions.get(session_id)
    if not session:
        return

    session["status"] = "running"
    print(f"[computer_use] Starting session {session_id}")

    try:
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        actions = session["actions"]

        # Build the instruction prompt for computer use
        instructions = _build_cu_instructions(actions, session["context"])

        # Use Claude's computer use with tool_use
        # The computer use loop: send screenshot → get action → execute → repeat
        messages = [{"role": "user", "content": instructions}]

        tools = [
            {
                "type": "computer_20250124",
                "name": "computer",
                "display_width_px": 1920,
                "display_height_px": 1080,
                "display_number": 1,
            }
        ]

        max_iterations = 30  # safety limit
        iteration = 0

        while iteration < max_iterations:
            iteration += 1
            print(f"[computer_use] Session {session_id} iteration {iteration}")

            response = client.messages.create(
                model=CU_MODEL,
                max_tokens=4096,
                tools=tools,
                messages=messages,
                betas=["computer-use-2025-01-24"],
            )

            # Check if Claude wants to use computer
            tool_uses = [b for b in response.content if b.type == "tool_use"]
            text_blocks = [b for b in response.content if b.type == "text"]

            if not tool_uses:
                # Claude is done — extract final text
                final_text = " ".join(b.text for b in text_blocks)
                session["result"] = {
                    "status": "completed",
                    "message": final_text,
                    "steps_taken": iteration,
                }
                session["status"] = "completed"
                print(f"[computer_use] Session {session_id} completed in {iteration} steps")
                break

            # Process each tool use (screenshot, click, type, etc.)
            tool_results = []
            for tool_use in tool_uses:
                step = {
                    "iteration": iteration,
                    "tool": tool_use.name,
                    "action": tool_use.input.get("action", ""),
                    "timestamp": time.time(),
                }
                session["steps"].append(step)

                # Execute the computer action
                result = await _execute_computer_action(tool_use)
                tool_results.append({
                    "type": "tool_result",
                    "tool_use_id": tool_use.id,
                    "content": result,
                })

            # Feed results back to Claude
            messages.append({"role": "assistant", "content": response.content})
            messages.append({"role": "user", "content": tool_results})

        if iteration >= max_iterations:
            session["status"] = "failed"
            session["error"] = "Max iterations reached"

    except Exception as e:
        session["status"] = "failed"
        session["error"] = str(e)
        print(f"[computer_use] Session {session_id} failed: {e}")


def _build_cu_instructions(actions: list, context: dict) -> str:
    """Build a natural language instruction for Claude computer use."""
    parts = [
        "You are controlling a Mac with Microsoft Excel open.",
        "The user's Excel workbook is already open and visible on screen.",
        "Perform the following tasks using Excel's GUI menus and dialogs:",
        "",
    ]

    for i, action in enumerate(actions, 1):
        action_type = action.get("type", "")
        payload = action.get("payload", {})

        if action_type == "create_data_table":
            sheet = payload.get("sheet", "")
            table_range = payload.get("range", "")
            row_input = payload.get("row_input_cell", "")
            col_input = payload.get("col_input_cell", "")
            parts.append(f"{i}. Navigate to sheet '{sheet}'")
            parts.append(f"   Select the range {table_range}")
            parts.append(f"   Go to Data tab → What-If Analysis → Data Table")
            if row_input and col_input:
                parts.append(f"   Set Row input cell: {row_input}")
                parts.append(f"   Set Column input cell: {col_input}")
                parts.append(f"   Click OK (this is a two-variable data table)")
            elif col_input:
                parts.append(f"   Leave Row input cell empty")
                parts.append(f"   Set Column input cell: {col_input}")
                parts.append(f"   Click OK (this is a one-variable data table)")
            elif row_input:
                parts.append(f"   Set Row input cell: {row_input}")
                parts.append(f"   Leave Column input cell empty")
                parts.append(f"   Click OK")

        elif action_type == "scenario_manager":
            parts.append(f"{i}. Go to Data tab → What-If Analysis → Scenario Manager")
            name = payload.get("name", "")
            cells = payload.get("changing_cells", "")
            values = payload.get("values", [])
            parts.append(f"   Click 'Add...' to create a new scenario")
            parts.append(f"   Scenario name: '{name}'")
            parts.append(f"   Changing cells: {cells}")
            parts.append(f"   Click OK, then enter values: {values}")
            parts.append(f"   Click OK, then Close")

        elif action_type == "save_solver_scenario":
            parts.append(f"{i}. Go to Data tab → Solver")
            parts.append(f"   Set up and run Solver with the current parameters")
            parts.append(f"   When results dialog appears, click 'Save Scenario...'")
            name = payload.get("name", "Solver Solution")
            parts.append(f"   Name it '{name}' and click OK")

        elif action_type == "run_solver":
            parts.append(f"{i}. Go to Data tab → Solver")
            objective = payload.get("objective_cell", "")
            goal = payload.get("goal", "max")
            changing = payload.get("changing_cells", "")
            constraints = payload.get("constraints", [])
            parts.append(f"   Set Objective: {objective}")
            parts.append(f"   To: {'Max' if goal == 'max' else 'Min' if goal == 'min' else 'Value Of: ' + str(goal)}")
            parts.append(f"   By Changing: {changing}")
            for c in constraints:
                parts.append(f"   Add constraint: {c}")
            parts.append(f"   Click Solve")

        elif action_type == "goal_seek":
            parts.append(f"{i}. Go to Data tab → What-If Analysis → Goal Seek")
            parts.append(f"   Set cell: {payload.get('set_cell', '')}")
            parts.append(f"   To value: {payload.get('to_value', '')}")
            parts.append(f"   By changing cell: {payload.get('changing_cell', '')}")
            parts.append(f"   Click OK")

        elif action_type == "scenario_summary":
            parts.append(f"{i}. Go to Data tab → What-If Analysis → Scenario Manager")
            parts.append(f"   Click 'Summary...'")
            result_cells = payload.get("result_cells", "")
            parts.append(f"   Result cells: {result_cells}")
            parts.append(f"   Report type: Scenario summary")
            parts.append(f"   Click OK")

        elif action_type == "run_toolpak":
            parts.append(f"{i}. Go to Data tab → Data Analysis (Analysis ToolPak)")
            tool = payload.get("tool", "Descriptive Statistics")
            parts.append(f"   Select: {tool}")
            input_range = payload.get("input_range", "")
            output_range = payload.get("output_range", "")
            parts.append(f"   Input Range: {input_range}")
            if output_range:
                parts.append(f"   Output Range: {output_range}")
            options = payload.get("options", {})
            if options.get("summary_statistics"):
                parts.append(f"   Check 'Summary statistics'")
            if options.get("labels_in_first_row"):
                parts.append(f"   Check 'Labels in first row'")
            parts.append(f"   Click OK")

        else:
            parts.append(f"{i}. {action_type}: {json.dumps(payload)}")

    parts.append("")
    parts.append("After completing all tasks, confirm what was done.")
    parts.append("IMPORTANT: Take a screenshot first to see the current state.")

    return "\n".join(parts)


async def _execute_computer_action(tool_use) -> list:
    """Execute a computer use action and return the result.

    NOTE: This is the platform-specific layer. On a real deployment,
    this would use pyautogui/subprocess to take screenshots, move mouse,
    click, and type. For now, it returns a placeholder that guides
    Claude through the interaction.

    In production, replace this with actual screen capture + input injection.
    """
    action = tool_use.input.get("action", "")
    print(f"[computer_use] Action: {action} | Input: {json.dumps(tool_use.input)[:200]}")

    if action == "screenshot":
        # In production: capture actual screenshot
        # For now: return a message
        return [{"type": "text", "text": "Screenshot captured. Excel is visible with the workbook open."}]

    elif action == "mouse_move":
        x = tool_use.input.get("coordinate", [0, 0])
        print(f"[computer_use] Mouse move to {x}")
        return [{"type": "text", "text": f"Mouse moved to ({x[0]}, {x[1]})"}]

    elif action == "left_click":
        x = tool_use.input.get("coordinate", [0, 0])
        print(f"[computer_use] Click at {x}")
        return [{"type": "text", "text": f"Clicked at ({x[0]}, {x[1]})"}]

    elif action == "type":
        text = tool_use.input.get("text", "")
        print(f"[computer_use] Type: {text[:50]}")
        return [{"type": "text", "text": f"Typed: {text}"}]

    elif action == "key":
        key = tool_use.input.get("text", "")
        print(f"[computer_use] Key: {key}")
        return [{"type": "text", "text": f"Pressed key: {key}"}]

    else:
        return [{"type": "text", "text": f"Action '{action}' executed."}]
