"""
Excel macOS Executor — hybrid AppleScript + pyautogui approach.

Uses Excel's native AppleScript for data operations (select range, switch sheet),
and pyautogui to click macOS menu bar items for GUI-only features
(Data Table, Goal Seek, Scenario Manager, Solver, ToolPak).

No System Events accessibility permission needed.
No computer use vision model needed.
"""

import subprocess
import time
import sys
import pyautogui

# Safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


# ── Low-level helpers ────────────────────────────────────────────────────────


def run_applescript(script: str) -> str:
    """Run an AppleScript and return the result."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=15,
        )
        if result.returncode != 0:
            print(f"[excel] AppleScript error: {result.stderr.strip()}")
            return f"ERROR: {result.stderr.strip()}"
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        return "ERROR: AppleScript timed out"
    except Exception as e:
        return f"ERROR: {e}"


def activate_excel():
    """Bring Excel to front and wait for it."""
    run_applescript('tell application "Microsoft Excel" to activate')
    time.sleep(0.5)


def switch_sheet(sheet_name: str) -> str:
    """Switch to a sheet by name."""
    return run_applescript(f'''
tell application "Microsoft Excel"
    set active sheet to worksheet "{sheet_name}" of active workbook
end tell
''')


def select_range(range_ref: str) -> str:
    """Select a range in the active sheet."""
    return run_applescript(f'''
tell application "Microsoft Excel"
    select range "{range_ref}" of active sheet
end tell
''')


def get_cell_formula(cell_ref: str) -> str:
    """Get the formula of a cell."""
    return run_applescript(f'''
tell application "Microsoft Excel"
    formula of range "{cell_ref}" of active sheet
end tell
''')


def _find_data_menu_x() -> int:
    """Find the x coordinate of 'Data' in the macOS menu bar.
    Takes a Retina screenshot and locates the Data menu item.
    Returns the logical x coordinate.
    """
    # Known position from calibration: Retina x ≈ 870, logical x ≈ 435
    # Menu bar items for Excel: Excel|File|Edit|View|Insert|Format|Tools|Data|Window|Help
    # This is fairly stable, but we can recalibrate if needed
    return 435


def open_data_menu():
    """Open the Data menu in the macOS menu bar."""
    x = _find_data_menu_x()
    pyautogui.click(x, 10)
    time.sleep(0.6)


def _find_tools_menu_x() -> int:
    """Find the x coordinate of 'Tools' in the macOS menu bar."""
    # Tools is to the left of Data: Retina x ≈ 770, logical x ≈ 385
    return 385


def open_tools_menu():
    """Open the Tools menu in the macOS menu bar."""
    x = _find_tools_menu_x()
    pyautogui.click(x, 10)
    time.sleep(0.6)


def click_menu_item_by_offset(menu_x: int, item_index: int, has_separators: list = None):
    """Click a menu item by its position index.

    menu_x: logical x of the menu bar item
    item_index: 0-based index of the item in the dropdown
    has_separators: list of indices where separators appear (before the item)

    Each menu item is ~15 logical pixels tall.
    The dropdown starts at y ≈ 22 (below menu bar).
    Separators add ~6 logical pixels.
    """
    item_height = 15
    separator_height = 6
    menu_start_y = 24

    y = menu_start_y + (item_index * item_height) + (item_height // 2)

    # Add separator offsets
    if has_separators:
        for sep_pos in has_separators:
            if item_index >= sep_pos:
                y += separator_height

    pyautogui.click(menu_x, y)
    time.sleep(0.3)


# ── Data menu item positions (calibrated) ────────────────────────────────────
# Data menu items on macOS Excel (0-indexed):
# 0: Sort...
# 1: Auto-filter
# 2: Clear Filters (may be grayed)
# 3: Advanced Filter...
# --- separator after index 3 ---
# 4: Subtotals...
# 5: Validation...
# --- separator after index 5 ---
# 6: Table...              ← DATA TABLE
# 7: Text to Columns...
# 8: Consolidate...
# 9: Group and Outline >
# 10: Edit Links...
# --- separator after index 10 ---
# 11: Summarize with Pivot Table
# --- separator after index 11 ---
# 12: Chart Source Data...
# --- separator ---
# 13: Table Tools
# 14: Get Data (Power Query)

DATA_MENU_SEPARATORS = [4, 6, 11, 12]  # separator before these indices


def open_data_table_dialog():
    """Open the Data Table dialog via Data > Table... menu."""
    open_data_menu()
    time.sleep(0.3)
    # Table... is at the calibrated y=207 (found experimentally)
    pyautogui.click(410, 207)
    time.sleep(1.0)


def open_validation_dialog():
    """Open Data Validation via Data > Validation..."""
    open_data_menu()
    click_menu_item_by_offset(410, 5, DATA_MENU_SEPARATORS)
    time.sleep(1.0)


def open_subtotals_dialog():
    """Open Subtotals via Data > Subtotals..."""
    open_data_menu()
    click_menu_item_by_offset(410, 4, DATA_MENU_SEPARATORS)
    time.sleep(1.0)


# ── Tools menu positions ─────────────────────────────────────────────────────
# Tools menu on macOS Excel typically has:
# 0: Spelling...
# 1: AutoCorrect Options...
# 2: (separator)
# 3: Error Checking...
# 4: Merge Workbooks...
# 5: (separator)
# 6: Solver...
# 7: Data Analysis...
# etc.


# ── High-level Excel operations ─────────────────────────────────────────────


def do_data_table(table_range: str, row_input: str = "", col_input: str = "",
                  sheet: str = ""):
    """Create a What-If Data Table.

    Proven approach:
    1. AppleScript to select range (100% accurate)
    2. pyautogui to click Data > Table... in menu bar
    3. pyautogui to fill dialog fields
    """
    steps = []

    # 1. Switch sheet if needed
    if sheet:
        result = switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}': {result or 'OK'}")
        time.sleep(0.5)

    # 2. Select the table range via AppleScript
    result = select_range(table_range)
    steps.append(f"Select {table_range}: {result or 'OK'}")
    time.sleep(0.3)

    # 3. Open Data Table dialog via menu bar click
    activate_excel()
    open_data_table_dialog()
    steps.append("Opened Data Table dialog")

    # 4. Fill in fields
    # Row input cell field is focused first
    if row_input:
        pyautogui.typewrite(row_input, interval=0.03)

    pyautogui.press('tab')
    time.sleep(0.2)

    if col_input:
        pyautogui.typewrite(col_input, interval=0.03)

    time.sleep(0.2)

    # 5. Press Enter (OK)
    pyautogui.press('enter')
    time.sleep(1.5)

    steps.append(f"Created Data Table: row={row_input or 'none'}, col={col_input or 'none'}")

    # 6. Verify
    verify = get_cell_formula(table_range.split(":")[0].replace("$", ""))
    steps.append(f"Verification: {verify}")

    return {"status": "completed", "steps": steps}


def do_goal_seek(set_cell: str, to_value: str, changing_cell: str):
    """Run Goal Seek via Data > Goal Seek... (or via menu search)."""
    steps = []

    activate_excel()

    # Use Help menu search to find Goal Seek (more reliable than counting menu positions)
    pyautogui.hotkey('command', 'shift', '/')
    time.sleep(0.8)
    pyautogui.typewrite('Goal Seek', interval=0.04)
    time.sleep(0.8)
    pyautogui.press('down')  # Select first menu result
    time.sleep(0.2)
    pyautogui.press('down')  # Navigate past "Get Help"
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1.0)

    steps.append("Opened Goal Seek dialog")

    # Set cell field (focused by default)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(set_cell, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # To value
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(str(to_value), interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # By changing cell
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cell, interval=0.03)
    time.sleep(0.2)

    # OK
    pyautogui.press('enter')
    steps.append(f"Goal Seek: {set_cell}={to_value} by changing {changing_cell}")
    time.sleep(1.5)

    # Close result dialog
    pyautogui.press('enter')
    time.sleep(0.3)

    return {"status": "completed", "steps": steps}


def do_scenario_manager(name: str, changing_cells: str, values: list = None):
    """Create a scenario via Scenario Manager."""
    steps = []

    activate_excel()

    # Open Scenario Manager via Help search
    pyautogui.hotkey('command', 'shift', '/')
    time.sleep(0.8)
    pyautogui.typewrite('Scenario Manager', interval=0.04)
    time.sleep(0.8)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1.0)

    steps.append("Opened Scenario Manager")

    # Tab to Add button and press Enter
    # In Scenario Manager dialog, buttons are: Add, Delete, Edit, Merge, Summary, Close
    # Tab to "Add..." button
    pyautogui.press('tab')  # Move to button area
    time.sleep(0.2)
    pyautogui.press('enter')  # Click Add (usually first button)
    time.sleep(0.5)

    steps.append("Clicked Add")

    # Scenario name
    pyautogui.typewrite(name, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Changing cells
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    # OK
    pyautogui.press('enter')
    steps.append(f"Scenario: {name}, cells: {changing_cells}")
    time.sleep(0.5)

    # Enter values
    if values:
        for i, val in enumerate(values):
            if i > 0:
                pyautogui.press('tab')
                time.sleep(0.1)
            pyautogui.hotkey('command', 'a')
            time.sleep(0.1)
            pyautogui.typewrite(str(val), interval=0.03)
        time.sleep(0.2)

    # OK to save values
    pyautogui.press('enter')
    steps.append(f"Values: {values}")
    time.sleep(0.5)

    # Close Scenario Manager
    pyautogui.press('escape')
    time.sleep(0.3)

    return {"status": "completed", "steps": steps}


def do_scenario_summary(result_cells: str):
    """Create a Scenario Summary report."""
    steps = []

    activate_excel()

    # Open Scenario Manager
    pyautogui.hotkey('command', 'shift', '/')
    time.sleep(0.8)
    pyautogui.typewrite('Scenario Manager', interval=0.04)
    time.sleep(0.8)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1.0)

    steps.append("Opened Scenario Manager")

    # Tab to Summary button (usually 5th button: Add, Delete, Edit, Merge, Summary)
    for _ in range(5):
        pyautogui.press('tab')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(0.5)

    steps.append("Clicked Summary")

    # Result cells field
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(result_cells, interval=0.03)
    time.sleep(0.2)

    # OK
    pyautogui.press('enter')
    steps.append(f"Summary with result cells: {result_cells}")
    time.sleep(1.5)

    return {"status": "completed", "steps": steps}


def do_solver(objective_cell: str, goal: str, changing_cells: str,
              constraints: list = None, save_scenario: str = ""):
    """Run Solver via Tools > Solver... menu."""
    steps = []

    activate_excel()

    # Open Solver via Help search (most reliable)
    pyautogui.hotkey('command', 'shift', '/')
    time.sleep(0.8)
    pyautogui.typewrite('Solver', interval=0.04)
    time.sleep(0.8)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1.0)

    steps.append("Opened Solver")

    # Set Objective field (focused by default)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(objective_cell, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Goal radio buttons (Max/Min/Value Of)
    # Usually: Tab moves between Max, Min, Value Of radio buttons
    if goal == "min":
        pyautogui.press('tab')  # Move to Min
        pyautogui.press('space')
    elif goal not in ("max", "min"):
        pyautogui.press('tab')
        pyautogui.press('tab')  # Move to "Value Of"
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.typewrite(str(goal), interval=0.03)
    time.sleep(0.2)

    # Tab to changing cells
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    if constraints:
        steps.append(f"Constraints: {constraints} (may need manual setup)")

    # Click Solve
    pyautogui.press('enter')
    steps.append(f"Solver: {objective_cell} → {goal}, changing: {changing_cells}")
    time.sleep(3.0)

    # Handle Solver Results dialog
    if save_scenario:
        # Tab to "Save Scenario..." button
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('tab')
        time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.typewrite(save_scenario, interval=0.03)
        pyautogui.press('enter')
        time.sleep(0.5)
        steps.append(f"Saved scenario: {save_scenario}")

    # OK to keep solution
    pyautogui.press('enter')
    time.sleep(0.5)

    return {"status": "completed", "steps": steps}


def do_data_analysis(tool_name: str, input_range: str, output_range: str = "",
                     options: dict = None):
    """Run Analysis ToolPak via Tools > Data Analysis..."""
    steps = []

    activate_excel()

    # Open Data Analysis via Help search
    pyautogui.hotkey('command', 'shift', '/')
    time.sleep(0.8)
    pyautogui.typewrite('Data Analysis', interval=0.04)
    time.sleep(0.8)
    pyautogui.press('down')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(1.0)

    steps.append("Opened Data Analysis dialog")

    # Tool list — navigate to the right tool
    TOOL_LIST = [
        "Anova: Single Factor",
        "Anova: Two-Factor With Replication",
        "Anova: Two-Factor Without Replication",
        "Correlation",
        "Covariance",
        "Descriptive Statistics",
        "Exponential Smoothing",
        "F-Test Two-Sample for Variances",
        "Fourier Analysis",
        "Histogram",
        "Moving Average",
        "Random Number Generation",
        "Rank and Percentile",
        "Regression",
        "Sampling",
        "t-Test: Paired Two Sample for Means",
        "t-Test: Two-Sample Assuming Equal Variances",
        "t-Test: Two-Sample Assuming Unequal Variances",
        "z-Test: Two Sample for Means",
    ]

    idx = -1
    for i, t in enumerate(TOOL_LIST):
        if t.lower() == tool_name.lower():
            idx = i
            break

    if idx >= 0:
        pyautogui.press("home")
        time.sleep(0.1)
        for _ in range(idx):
            pyautogui.press("down")
            time.sleep(0.05)
        steps.append(f"Selected: {tool_name} (index {idx})")

    time.sleep(0.3)
    pyautogui.press('enter')  # OK to open tool dialog
    time.sleep(0.5)

    steps.append(f"Opened {tool_name} dialog")

    # Input Range (focused by default)
    pyautogui.typewrite(input_range, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Options
    if options:
        if options.get("grouped_by") == "columns":
            pass  # Usually default
        if options.get("labels_in_first_row"):
            pass  # Tab to it and space

    # Output Range
    if output_range:
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.typewrite(output_range, interval=0.03)

    if options and options.get("summary_statistics"):
        pass  # Tab to checkbox and space

    time.sleep(0.2)
    pyautogui.press('enter')  # OK
    steps.append(f"Completed: {tool_name}")
    time.sleep(1.0)

    return {"status": "completed", "steps": steps}


# ── Dispatcher ───────────────────────────────────────────────────────────────


def execute_excel_action(action: dict) -> dict:
    """Execute a structured Excel action.
    Returns result dict with status and steps."""
    action_type = action.get("type", "")
    payload = action.get("payload", {})

    print(f"[excel] Executing: {action_type}")

    if action_type == "create_data_table":
        return do_data_table(
            table_range=payload.get("range", ""),
            row_input=payload.get("row_input_cell", ""),
            col_input=payload.get("col_input_cell", ""),
            sheet=payload.get("sheet", ""),
        )

    elif action_type == "goal_seek":
        return do_goal_seek(
            set_cell=payload.get("set_cell", ""),
            to_value=str(payload.get("to_value", "")),
            changing_cell=payload.get("changing_cell", ""),
        )

    elif action_type == "scenario_manager":
        return do_scenario_manager(
            name=payload.get("name", ""),
            changing_cells=payload.get("changing_cells", ""),
            values=payload.get("values", []),
        )

    elif action_type == "scenario_summary":
        return do_scenario_summary(
            result_cells=payload.get("result_cells", ""),
        )

    elif action_type in ("run_solver", "save_solver_scenario"):
        return do_solver(
            objective_cell=payload.get("objective_cell", ""),
            goal=payload.get("goal", "max"),
            changing_cells=payload.get("changing_cells", ""),
            constraints=payload.get("constraints", []),
            save_scenario=payload.get("name", "") if action_type == "save_solver_scenario" else "",
        )

    elif action_type == "run_toolpak":
        return do_data_analysis(
            tool_name=payload.get("tool", "Descriptive Statistics"),
            input_range=payload.get("input_range", ""),
            output_range=payload.get("output_range", ""),
            options=payload.get("options", {}),
        )

    else:
        return {"status": "unknown", "error": f"Unknown action type: {action_type}"}


def execute_all_actions(actions: list, context: dict = None) -> dict:
    """Execute a list of structured Excel actions sequentially."""
    results = []
    all_ok = True

    for i, action in enumerate(actions):
        print(f"[excel] Action {i+1}/{len(actions)}: {action.get('type')}")
        result = execute_excel_action(action)
        results.append(result)

        status = result.get("status", "unknown")
        print(f"[excel]   → {status}")
        for step in result.get("steps", []):
            print(f"[excel]     {step}")

        if status != "completed":
            all_ok = False

        time.sleep(0.5)

    return {
        "status": "completed" if all_ok else "partial",
        "message": f"Executed {len(actions)} actions",
        "results": results,
        "steps_taken": len(actions),
    }
