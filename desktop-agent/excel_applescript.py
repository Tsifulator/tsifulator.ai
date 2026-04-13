"""
Excel macOS Executor — hybrid AppleScript + pyautogui approach.

Uses Excel's native AppleScript for data operations (select range, switch sheet),
and pyautogui to click macOS menu bar items for GUI-only features
(Data Table, Goal Seek, Scenario Manager, Solver, ToolPak).

Every action verifies its dialog opened before typing.
A global stop flag can abort mid-action (checked between every step).
"""

import subprocess
import time
import sys
import threading
import pyautogui

# Safety
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1


# ── Stop mechanism ──────────────────────────────────────────────────────────
# A threading.Event that can be set from a background thread (cancel watcher)
# to interrupt automation mid-action.

_stop_event = threading.Event()


class StopAutomation(Exception):
    """Raised when the user cancels automation."""
    pass


def set_stop():
    """Signal all automation to stop immediately."""
    _stop_event.set()


def clear_stop():
    """Clear the stop flag before starting a new automation run."""
    _stop_event.clear()


def check_stop():
    """Call between every pyautogui step. Raises StopAutomation if user cancelled."""
    if _stop_event.is_set():
        raise StopAutomation("Stopped by user")


# ── Low-level helpers ────────────────────────────────────────────────────────


def run_applescript(script: str) -> str:
    """Run an AppleScript and return the result."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=30,
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


def get_cell_value(cell_ref: str) -> str:
    """Get the value of a cell."""
    return run_applescript(f'''
tell application "Microsoft Excel"
    value of range "{cell_ref}" of active sheet
end tell
''')


def dismiss_excel_dialogs():
    """Dismiss any open Excel error/alert dialogs by pressing Escape and Enter.
    This handles the 'selection isn't valid' and similar error popups."""
    # Try pressing Escape first (closes most dialogs)
    pyautogui.press('escape')
    time.sleep(0.3)
    # Also try Enter (for OK-only dialogs)
    pyautogui.press('enter')
    time.sleep(0.3)
    # One more Escape for good measure
    pyautogui.press('escape')
    time.sleep(0.2)


def ensure_clean_state():
    """Make sure Excel is in a clean, predictable state before starting an action.
    Dismisses dialogs, cancels editing, selects A1."""
    activate_excel()
    time.sleep(0.3)
    # Dismiss any open dialogs/menus/cell editing
    for _ in range(4):
        pyautogui.press('escape')
        time.sleep(0.15)
    time.sleep(0.3)
    # Select A1 to ensure we're not in an unexpected cell
    run_applescript('tell application "Microsoft Excel" to select range "A1" of active sheet')
    time.sleep(0.2)


def verify_dialog(expected_title_fragment: str, timeout: float = 3.0) -> bool:
    """Check if a dialog/window containing expected_title_fragment is open.
    Uses AppleScript to inspect the front window name."""
    start = time.time()
    while time.time() - start < timeout:
        result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        try
            set winName to name of front window
            return winName
        on error
            return "NO_WINDOW"
        end try
    end tell
end tell
''')
        if expected_title_fragment.lower() in result.lower():
            print(f"[excel] Dialog verified: '{result}' matches '{expected_title_fragment}'")
            return True
        time.sleep(0.3)
    print(f"[excel] Dialog NOT found: expected '{expected_title_fragment}', got '{result}'")
    return False


# ── Direct AppleScript menu paths for each dialog ──────────────────────────
# These are much more reliable than Help menu search.

MENU_PATHS = {
    "Scenario Manager": '''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Scenario Manager…" of menu 1 of menu item "What-If Analysis" of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''',
    "Solver": '''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Solver…" of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''',
    "Data Analysis": '''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Data Analysis…" of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''',
    "Excel Add-ins": '''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Excel Add-ins…" of menu 1 of menu bar item "Tools" of menu bar 1
    end tell
end tell
''',
    "Goal Seek": '''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Goal Seek…" of menu 1 of menu item "What-If Analysis" of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''',
}

# What the front window title should contain when each dialog is open
DIALOG_TITLES = {
    "Scenario Manager": "Scenario Manager",
    "Solver": "Solver",
    "Data Analysis": "Data Analysis",
    "Excel Add-ins": "Add-Ins",
    "Goal Seek": "Goal Seek",
}


def open_dialog(name: str) -> bool:
    """Open a named dialog using direct AppleScript menu click.
    Returns True if the dialog was verified open, False if it failed.

    Tries: 1) direct menu click  2) alternative menu names  3) Help search
    Each attempt is verified with verify_dialog().
    """
    check_stop()

    activate_excel()
    time.sleep(0.3)

    expected_title = DIALOG_TITLES.get(name, name)

    # Attempt 1: Direct AppleScript menu click
    script = MENU_PATHS.get(name, "")
    if script:
        result = run_applescript(script)
        if "ERROR" not in result:
            time.sleep(1.0)
            if verify_dialog(expected_title):
                return True
            print(f"[excel] Direct menu click succeeded but dialog not verified")

    check_stop()

    # Attempt 2: Try without unicode ellipsis (use ... instead of …)
    if script:
        alt_script = script.replace("…", "...")
        if alt_script != script:
            result = run_applescript(alt_script)
            if "ERROR" not in result:
                time.sleep(1.0)
                if verify_dialog(expected_title):
                    return True

    check_stop()

    # Attempt 3: Help menu search (last resort)
    print(f"[excel] Falling back to Help menu search for '{name}'")
    activate_excel()
    time.sleep(0.3)
    for _ in range(3):
        pyautogui.press('escape')
        time.sleep(0.15)
    time.sleep(0.3)

    result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu bar item "Help" of menu bar 1
    end tell
end tell
''')
    time.sleep(1.0)

    if "ERROR" in result:
        pyautogui.hotkey('command', 'shift', '/')
        time.sleep(1.5)

    check_stop()
    pyautogui.typewrite(name, interval=0.04)
    time.sleep(1.0)
    pyautogui.press('enter')
    time.sleep(1.5)

    if verify_dialog(expected_title):
        return True

    print(f"[excel] FAILED to open dialog '{name}' after 3 attempts")
    return False


# ── Menu bar coordinates (calibrated for macOS Excel) ────────────────────────
# Menu bar: Apple | Excel | File | Edit | View | Insert | Format | Tools | Data | Window | Help
# These are LOGICAL coordinates (screen points, not Retina pixels)

MENU_X = {
    "tools": 385,
    "data": 435,
}

# Data > Table... is at calibrated y=207 (found experimentally)
DATA_TABLE_Y = 207


def open_data_table_dialog():
    """Open the Data Table dialog via Data > Table... (item 9 in Data menu).
    Uses menu item index for reliability — name matching fails due to ellipsis encoding."""
    result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item 9 of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''')
    if "ERROR" in result:
        print(f"[excel] Menu index click failed ({result}), trying by name...")
        result2 = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "Table..." of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''')
        if "ERROR" in result2:
            print(f"[excel] Name click also failed, trying Help search...")
            search_menu("Table", wait=1.5)
    time.sleep(1.0)


# ── High-level Excel operations ─────────────────────────────────────────────


def do_data_table(table_range: str, row_input: str = "", col_input: str = "",
                  sheet: str = ""):
    """Create a What-If Data Table."""
    steps = []
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.5)

    check_stop()

    # Select the table range via AppleScript (always accurate)
    select_range(table_range)
    steps.append(f"Selected {table_range}")
    time.sleep(0.3)

    check_stop()

    # Open Data Table dialog via menu bar click + verify
    activate_excel()
    open_data_table_dialog()
    time.sleep(0.5)
    if not verify_dialog("Table"):
        steps.append("FAILED: Data Table dialog did not open")
        return {"status": "failed", "error": "Dialog did not open", "steps": steps}
    steps.append("Data Table dialog open")

    check_stop()

    # Fill in fields — Row input cell field is focused first
    if row_input:
        pyautogui.typewrite(row_input, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)
    if col_input:
        pyautogui.typewrite(col_input, interval=0.03)
    time.sleep(0.2)

    check_stop()

    # OK
    pyautogui.press('enter')
    time.sleep(2.0)
    steps.append(f"Data Table: row={row_input or 'none'}, col={col_input or 'none'}")

    # Verify TABLE formula was created
    try:
        parts = table_range.replace("$", "").split(":")
        start_cell = parts[0]
        col = ''.join(c for c in start_cell if c.isalpha())
        row = int(''.join(c for c in start_cell if c.isdigit()))
        next_col = chr(ord(col[0]) + 1) if len(col) == 1 else col
        verify_cell = f"{next_col}{row + 1}"
        verify = get_cell_formula(verify_cell)
        if "TABLE" in verify.upper():
            steps.append("TABLE formula confirmed")
        else:
            steps.append(f"WARNING: TABLE not found in {verify_cell}: {verify}")
    except Exception as e:
        steps.append(f"Verification skipped: {e}")

    return {"status": "completed", "steps": steps}


def do_goal_seek(set_cell: str, to_value: str, changing_cell: str,
                 sheet: str = ""):
    """Run Goal Seek."""
    steps = []
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    check_stop()
    if not open_dialog("Goal Seek"):
        return {"status": "failed", "error": "Goal Seek dialog did not open", "steps": steps}
    steps.append("Goal Seek dialog open")

    check_stop()
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(set_cell, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    check_stop()
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(str(to_value), interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    check_stop()
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cell, interval=0.03)
    time.sleep(0.2)

    pyautogui.press('enter')
    steps.append(f"Goal Seek: {set_cell}={to_value} by changing {changing_cell}")
    time.sleep(2.0)
    pyautogui.press('enter')
    time.sleep(0.3)

    return {"status": "completed", "steps": steps}


def do_scenario_manager(name: str, changing_cells: str, values: list = None,
                        sheet: str = ""):
    """Create a scenario via Scenario Manager."""
    steps = []
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    check_stop()
    if not open_dialog("Scenario Manager"):
        return {"status": "failed", "error": "Scenario Manager did not open", "steps": steps}
    steps.append("Scenario Manager open")

    check_stop()
    # Click Add button
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.5)

    check_stop()
    # Scenario name
    pyautogui.typewrite(name, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Changing cells
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    check_stop()
    # OK → Scenario Values dialog
    pyautogui.press('enter')
    time.sleep(0.5)

    # Enter values if provided
    if values:
        for i, val in enumerate(values):
            if i > 0:
                pyautogui.press('tab')
                time.sleep(0.1)
            pyautogui.hotkey('command', 'a')
            time.sleep(0.1)
            pyautogui.typewrite(str(val), interval=0.03)
        time.sleep(0.2)

    check_stop()
    # OK to save values
    pyautogui.press('enter')
    steps.append(f"Scenario '{name}': cells={changing_cells}, values={values}")
    time.sleep(0.5)

    # Close Scenario Manager
    pyautogui.press('escape')
    time.sleep(0.3)
    pyautogui.press('escape')
    time.sleep(0.5)

    return {"status": "completed", "steps": steps}


def do_scenario_summary(result_cells: str, sheet: str = ""):
    """Create a Scenario Summary report."""
    steps = []
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    check_stop()
    if not open_dialog("Scenario Manager"):
        return {"status": "failed", "error": "Scenario Manager did not open", "steps": steps}
    steps.append("Scenario Manager open")

    check_stop()
    # Tab to Summary button (list, Add, Delete, Edit, Merge, Summary)
    for _ in range(5):
        pyautogui.press('tab')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(0.5)

    check_stop()
    # Summary dialog: Report type then Result cells
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(result_cells, interval=0.03)
    time.sleep(0.2)

    pyautogui.press('enter')
    steps.append(f"Summary: result cells = {result_cells}")
    time.sleep(2.0)

    return {"status": "completed", "steps": steps}


def do_solver(objective_cell: str, goal: str, changing_cells: str,
              constraints: list = None, save_scenario: str = "",
              restore_original: bool = True, sheet: str = ""):
    """Run Solver with full constraint support."""
    steps = []
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    check_stop()
    if not open_dialog("Solver"):
        return {"status": "failed", "error": "Solver did not open", "steps": steps}
    steps.append("Solver open")

    check_stop()
    # Set Objective
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(objective_cell, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Goal radio: Max (default) | Min | Value Of
    if goal == "min":
        pyautogui.press('tab')
        pyautogui.press('space')
    elif goal not in ("max", "min"):
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.typewrite(str(goal), interval=0.03)
    time.sleep(0.2)

    check_stop()
    # By Changing Variable Cells
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    # Add constraints
    if constraints:
        for i, constraint in enumerate(constraints):
            check_stop()
            cell_ref = constraint.get("cell", "")
            operator = constraint.get("operator", "<=")
            value = str(constraint.get("value", ""))

            # Tab to Add button
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('enter')
            time.sleep(0.5)

            # Cell Reference
            pyautogui.typewrite(cell_ref, interval=0.03)
            pyautogui.press('tab')
            time.sleep(0.2)

            # Operator dropdown
            operator_map = {"<=": 0, ">=": 1, "=": 2, "int": 3, "bin": 4}
            clicks = operator_map.get(operator, 0)
            for _ in range(clicks):
                pyautogui.press('down')
                time.sleep(0.1)

            pyautogui.press('tab')
            time.sleep(0.2)

            # Value (not needed for int/bin)
            if operator not in ("int", "bin"):
                pyautogui.typewrite(value, interval=0.03)
            time.sleep(0.2)

            if i < len(constraints) - 1:
                pyautogui.press('tab')
                time.sleep(0.1)
                pyautogui.press('enter')  # Add another
                time.sleep(0.5)
            else:
                pyautogui.press('enter')  # OK (done adding)
                time.sleep(0.5)

            steps.append(f"Constraint: {cell_ref} {operator} {value}")

    check_stop()
    # Solve
    pyautogui.press('enter')
    steps.append(f"Solver: {objective_cell} -> {goal}, changing: {changing_cells}")
    time.sleep(3.0)

    # Solver Results dialog
    if restore_original:
        pyautogui.press('tab')
        time.sleep(0.1)
        pyautogui.press('space')
        time.sleep(0.2)
        steps.append("Restore Original Values")

    if save_scenario:
        check_stop()
        tab_count = 2 if restore_original else 3
        for _ in range(tab_count):
            pyautogui.press('tab')
            time.sleep(0.1)
        pyautogui.press('enter')  # Save Scenario
        time.sleep(0.5)

        pyautogui.typewrite(save_scenario, interval=0.03)
        pyautogui.press('enter')
        time.sleep(0.5)
        steps.append(f"Saved scenario: {save_scenario}")

    # OK to close
    pyautogui.press('enter')
    time.sleep(0.5)

    return {"status": "completed", "steps": steps}


def do_data_analysis(tool_name: str, input_range: str, output_range: str = "",
                     options: dict = None, sheet: str = ""):
    """Run Analysis ToolPak tool with full option support."""
    steps = []
    if not options:
        options = {}

    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    check_stop()
    if not open_dialog("Data Analysis"):
        return {"status": "failed", "error": "Data Analysis dialog did not open", "steps": steps}
    steps.append("Data Analysis dialog open")

    # Tool list — navigate to the right tool by index
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
        # Navigate to top of list then arrow down to the tool
        pyautogui.press("home")
        time.sleep(0.1)
        for _ in range(idx):
            pyautogui.press("down")
            time.sleep(0.05)
        steps.append(f"Selected: {tool_name} (index {idx})")
    else:
        # Try partial match
        for i, t in enumerate(TOOL_LIST):
            if tool_name.lower() in t.lower():
                pyautogui.press("home")
                time.sleep(0.1)
                for _ in range(i):
                    pyautogui.press("down")
                    time.sleep(0.05)
                steps.append(f"Selected: {t} (partial match, index {i})")
                idx = i
                break

    time.sleep(0.3)
    pyautogui.press('enter')  # OK to open tool dialog
    time.sleep(0.8)
    steps.append(f"Opened {tool_name} dialog")

    # ── Descriptive Statistics dialog layout (most common):
    # Input Range: [field]
    # Grouped By: (Columns) (Rows) — radio buttons
    # [x] Labels in First Row — checkbox
    # Output options: (Output Range) (New Worksheet Ply) (New Workbook) — radio
    # Output Range: [field]
    # [x] Summary Statistics — checkbox
    # [x] Confidence Level for Mean: [field]
    # [x] Kth Largest: [field]
    # [x] Kth Smallest: [field]

    # Input Range (focused by default)
    pyautogui.typewrite(input_range, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Grouped By radio buttons (Columns is default)
    if options.get("grouped_by") == "rows":
        # Tab to Rows radio and select it
        pyautogui.press('tab')  # Move to Rows radio
        pyautogui.press('space')
        time.sleep(0.1)
    else:
        pyautogui.press('tab')  # Skip past Grouped By
    time.sleep(0.1)

    # Labels in First Row checkbox
    pyautogui.press('tab')
    if options.get("labels_in_first_row"):
        pyautogui.press('space')  # Check the box
    time.sleep(0.1)

    # Output options — Tab to Output Range radio
    pyautogui.press('tab')  # Output Range radio (usually selected by default)
    time.sleep(0.1)

    # Output Range field
    pyautogui.press('tab')
    if output_range:
        pyautogui.typewrite(output_range, interval=0.03)
    time.sleep(0.1)

    # For Descriptive Statistics, handle additional options
    if tool_name.lower() == "descriptive statistics":
        # Tab past New Worksheet Ply and New Workbook radios
        pyautogui.press('tab')  # New Worksheet Ply
        pyautogui.press('tab')  # New Workbook
        time.sleep(0.1)

        # Summary Statistics checkbox
        pyautogui.press('tab')
        if options.get("summary_statistics", True):  # Default to True for Descriptive Stats
            pyautogui.press('space')
        time.sleep(0.1)

        # Confidence Level checkbox + field
        pyautogui.press('tab')
        if options.get("confidence_level"):
            pyautogui.press('space')
            pyautogui.press('tab')
            pyautogui.hotkey('command', 'a')
            conf = options.get("confidence_level")
            if isinstance(conf, (int, float)):
                pyautogui.typewrite(str(conf), interval=0.03)
        else:
            pyautogui.press('tab')  # Skip confidence field
        time.sleep(0.1)

        # Kth Largest
        pyautogui.press('tab')
        if options.get("kth_largest"):
            pyautogui.press('space')
            pyautogui.press('tab')
            pyautogui.hotkey('command', 'a')
            pyautogui.typewrite(str(options["kth_largest"]), interval=0.03)
        else:
            pyautogui.press('tab')
        time.sleep(0.1)

        # Kth Smallest
        pyautogui.press('tab')
        if options.get("kth_smallest"):
            pyautogui.press('space')
            pyautogui.press('tab')
            pyautogui.hotkey('command', 'a')
            pyautogui.typewrite(str(options["kth_smallest"]), interval=0.03)
        else:
            pyautogui.press('tab')
        time.sleep(0.1)

    time.sleep(0.2)
    pyautogui.press('enter')  # OK — run the analysis
    steps.append(f"Completed: {tool_name}")
    time.sleep(1.5)

    # Dismiss any error dialogs
    pyautogui.press('enter')
    time.sleep(0.3)

    return {"status": "completed", "steps": steps}


def do_install_addins(addins: list):
    """Install Excel add-ins (Solver, Analysis ToolPak) via Tools > Excel Add-ins."""
    steps = []
    ensure_clean_state()

    check_stop()
    if not open_dialog("Excel Add-ins"):
        return {"status": "failed", "error": "Excel Add-ins dialog did not open", "steps": steps}
    steps.append("Excel Add-ins dialog open")

    # The add-ins list order:
    # 1. Analysis ToolPak
    # 2. Analysis ToolPak - VBA
    # 3. Solver Add-in
    # Check/uncheck as needed

    want_toolpak = any("ToolPak" in a or "toolpak" in a.lower() for a in addins)
    want_solver = any("Solver" in a or "solver" in a.lower() for a in addins)

    if want_toolpak:
        pyautogui.press('space')  # Toggle Analysis ToolPak
        steps.append("Toggled Analysis ToolPak")
    pyautogui.press('down')
    time.sleep(0.1)
    pyautogui.press('down')  # Skip VBA
    time.sleep(0.1)
    if want_solver:
        pyautogui.press('space')  # Toggle Solver
        steps.append("Toggled Solver Add-in")

    time.sleep(0.2)
    pyautogui.press('enter')  # OK
    time.sleep(3.0)

    # Aggressively dismiss any followup dialogs and menus
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('escape')
    time.sleep(0.3)
    pyautogui.press('escape')
    time.sleep(0.3)
    pyautogui.press('escape')
    time.sleep(0.5)

    # Click on the spreadsheet area to ensure focus is back on the sheet
    activate_excel()
    time.sleep(0.5)
    # Click a safe cell to make sure we're out of any dialog/menu state
    run_applescript('tell application "Microsoft Excel" to select range "A1" of active sheet')
    time.sleep(0.5)

    return {"status": "completed", "steps": steps}


def do_uninstall_addins(addins: list):
    """Uninstall Excel add-ins via Tools > Excel Add-ins."""
    # Same as install — toggling unchecks them
    return do_install_addins(addins)


# ── Dispatcher ───────────────────────────────────────────────────────────────


def execute_excel_action(action: dict) -> dict:
    """Execute a structured Excel action.
    Returns result dict with status and steps."""
    action_type = action.get("type", "")
    payload = action.get("payload", {})

    print(f"[excel] Executing: {action_type}")

    try:
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
                sheet=payload.get("sheet", ""),
            )

        elif action_type == "scenario_manager":
            return do_scenario_manager(
                name=payload.get("name", ""),
                changing_cells=payload.get("changing_cells", ""),
                values=payload.get("values", []),
                sheet=payload.get("sheet", ""),
            )

        elif action_type == "scenario_summary":
            return do_scenario_summary(
                result_cells=payload.get("result_cells", ""),
                sheet=payload.get("sheet", ""),
            )

        elif action_type in ("run_solver", "save_solver_scenario"):
            return do_solver(
                objective_cell=payload.get("objective_cell", ""),
                goal=payload.get("goal", "max"),
                changing_cells=payload.get("changing_cells", ""),
                constraints=payload.get("constraints", []),
                save_scenario=payload.get("name", "") if action_type == "save_solver_scenario" else "",
                restore_original=payload.get("restore_original", True),
                sheet=payload.get("sheet", ""),
            )

        elif action_type == "run_toolpak":
            return do_data_analysis(
                tool_name=payload.get("tool", "Descriptive Statistics"),
                input_range=payload.get("input_range", ""),
                output_range=payload.get("output_range", ""),
                options=payload.get("options", {}),
                sheet=payload.get("sheet", ""),
            )

        elif action_type == "install_addins":
            return do_install_addins(payload.get("addins", []))

        elif action_type == "uninstall_addins":
            return do_uninstall_addins(payload.get("addins", []))

        else:
            return {"status": "unknown", "error": f"Unknown action type: {action_type}"}

    except StopAutomation:
        # User pressed Stop — clean up and propagate
        print(f"[excel] STOPPED by user during {action_type}")
        try:
            _emergency_cleanup()
        except Exception:
            pass
        return {"status": "cancelled", "error": "Stopped by user", "steps": [f"Stopped during {action_type}"]}

    except Exception as e:
        # Catch any unexpected errors, clean up, and report
        print(f"[excel] ERROR in {action_type}: {e}")
        try:
            _emergency_cleanup()
        except Exception:
            pass
        return {"status": "failed", "error": str(e), "steps": [f"Exception: {e}"]}


def execute_all_actions(actions: list, context: dict = None, cancel_check=None) -> dict:
    """Execute a list of structured Excel actions sequentially.

    cancel_check: optional callable returning True if the user pressed Stop.
    """
    results = []
    all_ok = True
    cancelled = False

    for i, action in enumerate(actions):
        # Check for user cancellation before each action (via HTTP)
        if cancel_check and cancel_check():
            print(f"[excel] CANCELLED by user before action {i+1}/{len(actions)}")
            cancelled = True
            _emergency_cleanup()
            break

        # Also check via the global stop flag (set by background watcher)
        if _stop_event.is_set():
            print(f"[excel] STOP FLAG set before action {i+1}/{len(actions)}")
            cancelled = True
            _emergency_cleanup()
            break

        print(f"[excel] Action {i+1}/{len(actions)}: {action.get('type')}")
        result = execute_excel_action(action)
        results.append(result)

        status = result.get("status", "unknown")
        print(f"[excel]   -> {status}")
        for step in result.get("steps", []):
            print(f"[excel]     {step}")

        # If the action was cancelled mid-execution, stop the whole sequence
        if status == "cancelled":
            cancelled = True
            break

        if status not in ("completed",):
            all_ok = False

        time.sleep(0.5)

    if cancelled:
        return {
            "status": "cancelled",
            "message": f"Stopped by user after {len(results)}/{len(actions)} actions",
            "results": results,
            "steps_taken": len(results),
        }

    return {
        "status": "completed" if all_ok else "partial",
        "message": f"Executed {len(actions)} actions ({sum(1 for r in results if r.get('status') == 'completed')}/{len(actions)} OK)",
        "results": results,
        "steps_taken": len(actions),
    }


def _emergency_cleanup():
    """Dismiss any open dialogs and return Excel to a safe state."""
    print("[excel] Emergency cleanup — dismissing dialogs...")
    for _ in range(5):
        pyautogui.press('escape')
        time.sleep(0.2)
    # Try to select A1 to get back to a known state
    try:
        activate_excel()
        time.sleep(0.3)
        run_applescript('tell application "Microsoft Excel" to select range "A1" of active sheet')
    except Exception:
        pass
    print("[excel] Cleanup done")
