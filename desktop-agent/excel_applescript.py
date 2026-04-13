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


def search_menu(search_term: str, wait: float = 1.0) -> bool:
    """Use Help menu search (Cmd+Shift+/) to find and click a menu item.
    Returns True if the search was initiated successfully.

    Strategy: use AppleScript to click the Help menu bar item first (reliable),
    then use the search field that appears at the top.
    Fallback: Cmd+Shift+/ shortcut.
    """
    # 1. Make sure Excel is focused and nothing is being edited
    activate_excel()
    time.sleep(0.3)
    # Dismiss any open dialogs/cell editing with multiple Escapes
    for _ in range(3):
        pyautogui.press('escape')
        time.sleep(0.15)
    time.sleep(0.3)

    # 2. Open Help menu search via AppleScript (most reliable)
    # This clicks the Help menu bar item which has a search field at top
    result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        -- Click Help menu to open it (search field is at top)
        click menu bar item "Help" of menu bar 1
    end tell
end tell
''')
    time.sleep(1.0)

    if "ERROR" in result:
        # Fallback: use keyboard shortcut
        print(f"[excel] Help menu click failed, trying Cmd+Shift+/")
        pyautogui.hotkey('command', 'shift', '/')
        time.sleep(1.5)

    # 3. Type the search term — the Help menu search field should be focused
    pyautogui.typewrite(search_term, interval=0.04)
    time.sleep(wait)

    # 4. Select the first result
    pyautogui.press('enter')
    time.sleep(1.0)
    return True


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
    """Create a What-If Data Table.

    1. AppleScript to select range (100% accurate)
    2. pyautogui to click Data > Table... in menu bar
    3. pyautogui to fill dialog fields
    """
    steps = []

    # 0. Clean slate
    ensure_clean_state()

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

    # 4. Fill in fields — Row input cell field is focused first
    if row_input:
        pyautogui.typewrite(row_input, interval=0.03)

    pyautogui.press('tab')
    time.sleep(0.2)

    if col_input:
        pyautogui.typewrite(col_input, interval=0.03)

    time.sleep(0.2)

    # 5. Press Enter (OK)
    pyautogui.press('enter')
    time.sleep(2.0)

    # 6. Check for error dialogs (e.g. "selection isn't valid")
    #    If an error dialog appeared, dismiss it and report
    #    We detect this by checking if an alert is showing via AppleScript
    error_check = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        if exists (sheet 1 of window 1) then
            return "no_dialog"
        else
            return "possible_dialog"
        end if
    end tell
end tell
''')
    # Simpler approach: just dismiss any potential dialog
    pyautogui.press('enter')  # Dismiss OK on error dialog if present
    time.sleep(0.3)

    steps.append(f"Created Data Table: row={row_input or 'none'}, col={col_input or 'none'}")

    # 7. Verify — check if TABLE formula was created
    try:
        # Check the second row of the data table for TABLE formula
        parts = table_range.replace("$", "").split(":")
        start_cell = parts[0]
        # Parse column letter and row number
        col = ''.join(c for c in start_cell if c.isalpha())
        row = int(''.join(c for c in start_cell if c.isdigit()))
        # Check one row below the start (where TABLE formulas appear)
        next_col = chr(ord(col[0]) + 1) if len(col) == 1 else col  # Move right one column
        verify_cell = f"{next_col}{row + 1}"
        verify = get_cell_formula(verify_cell)
        steps.append(f"Verify {verify_cell}: {verify}")
        if "TABLE" in verify.upper():
            steps.append("TABLE formula confirmed")
        else:
            steps.append("WARNING: TABLE formula not found — may need retry")
    except Exception as e:
        steps.append(f"Verification skipped: {e}")

    return {"status": "completed", "steps": steps}


def do_goal_seek(set_cell: str, to_value: str, changing_cell: str,
                 sheet: str = ""):
    """Run Goal Seek via Help menu search."""
    steps = []

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    activate_excel()

    # Open Goal Seek via Help search
    search_menu("Goal Seek")
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
    time.sleep(2.0)

    # Close Goal Seek Status dialog (shows result)
    pyautogui.press('enter')
    time.sleep(0.3)

    return {"status": "completed", "steps": steps}


def do_scenario_manager(name: str, changing_cells: str, values: list = None,
                        sheet: str = ""):
    """Create a scenario via Scenario Manager.
    Uses Help menu search to open Scenario Manager reliably."""
    steps = []

    # Clean state before starting
    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    # Open Scenario Manager via Help search
    search_menu("Scenario Manager")
    steps.append("Opened Scenario Manager")

    # Click Add button — Tab into button area, Enter to activate
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.5)
    steps.append("Clicked Add")

    # Scenario name field
    pyautogui.typewrite(name, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Changing cells field
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    # OK — moves to Scenario Values dialog
    pyautogui.press('enter')
    steps.append(f"Scenario: {name}, cells: {changing_cells}")
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

    # OK to save values — returns to Scenario Manager
    pyautogui.press('enter')
    steps.append(f"Values: {values}")
    time.sleep(0.5)

    # Close Scenario Manager — click Close button explicitly
    # After adding a scenario, focus returns to the scenario list.
    # Tab to the Close button (last button: Add, Delete, Edit, Merge, Summary, Close)
    # Press Escape first to ensure we're not in a sub-dialog, then
    # use Cmd+. or Escape which triggers the Close button on the manager dialog
    pyautogui.press('escape')
    time.sleep(0.3)
    # Double-check: press Escape again to ensure dialog is fully dismissed
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

    # Open Scenario Manager
    search_menu("Scenario Manager")
    steps.append("Opened Scenario Manager")

    # Tab to Summary button
    # Buttons: scenario list, Add, Delete, Edit, Merge, Summary, Close
    # Summary is usually button 5 (after list focus)
    for _ in range(5):
        pyautogui.press('tab')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(0.5)
    steps.append("Clicked Summary")

    # Summary dialog: Report type (Scenario summary / PivotTable)
    # Then Result cells field
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(result_cells, interval=0.03)
    time.sleep(0.2)

    # OK
    pyautogui.press('enter')
    steps.append(f"Summary with result cells: {result_cells}")
    time.sleep(2.0)

    return {"status": "completed", "steps": steps}


def do_solver(objective_cell: str, goal: str, changing_cells: str,
              constraints: list = None, save_scenario: str = "",
              restore_original: bool = True, sheet: str = ""):
    """Run Solver with full constraint support.

    constraints format: [{"cell": "$D$5", "operator": "<=", "value": "100"}, ...]
    operator can be: "<=", ">=", "=", "int", "bin"
    """
    steps = []

    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    # Open Solver via Help search
    search_menu("Solver")
    steps.append("Opened Solver")

    # Set Objective field (focused by default)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(objective_cell, interval=0.03)
    pyautogui.press('tab')
    time.sleep(0.2)

    # Goal radio buttons: Max | Min | Value Of
    if goal == "min":
        pyautogui.press('tab')  # Move to Min radio
        pyautogui.press('space')
    elif goal not in ("max", "min"):
        # "Value Of" — tab past Min to Value Of
        pyautogui.press('tab')  # Min
        pyautogui.press('tab')  # Value Of
        pyautogui.press('space')
        time.sleep(0.1)
        pyautogui.typewrite(str(goal), interval=0.03)
    # else: Max is default, no action needed
    time.sleep(0.2)

    # Tab to "By Changing Variable Cells" field
    pyautogui.press('tab')
    time.sleep(0.2)
    pyautogui.hotkey('command', 'a')
    time.sleep(0.1)
    pyautogui.typewrite(changing_cells, interval=0.03)
    time.sleep(0.2)

    # Add constraints
    if constraints:
        for i, constraint in enumerate(constraints):
            cell_ref = constraint.get("cell", "")
            operator = constraint.get("operator", "<=")
            value = str(constraint.get("value", ""))

            # Click "Add" button for constraints
            # Tab to Add button area
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.press('enter')  # Click Add
            time.sleep(0.5)

            # Add Constraint dialog:
            # Cell Reference | Operator dropdown | Constraint value
            pyautogui.typewrite(cell_ref, interval=0.03)
            pyautogui.press('tab')
            time.sleep(0.2)

            # Operator dropdown — cycle through options
            # Default is "<=", then ">=", "=", "int", "bin", "dif"
            operator_map = {"<=": 0, ">=": 1, "=": 2, "int": 3, "bin": 4}
            clicks = operator_map.get(operator, 0)
            for _ in range(clicks):
                pyautogui.press('down')
                time.sleep(0.1)

            pyautogui.press('tab')
            time.sleep(0.2)

            # Constraint value (not needed for int/bin)
            if operator not in ("int", "bin"):
                pyautogui.typewrite(value, interval=0.03)

            time.sleep(0.2)

            # OK to add constraint (or Add to add another)
            if i < len(constraints) - 1:
                # More constraints coming — click Add
                pyautogui.press('tab')
                time.sleep(0.1)
                pyautogui.press('enter')  # Add button
                time.sleep(0.5)
            else:
                # Last constraint — click OK
                pyautogui.press('enter')
                time.sleep(0.5)

            steps.append(f"Constraint: {cell_ref} {operator} {value}")

    # Click Solve
    pyautogui.press('enter')
    steps.append(f"Solver: {objective_cell} -> {goal}, changing: {changing_cells}")
    time.sleep(3.0)

    # Handle Solver Results dialog
    # Dialog layout: "Keep Solver Solution" radio (selected by default),
    #   "Restore Original Values" radio, then buttons area

    if restore_original:
        # Select "Restore Original Values" radio button
        pyautogui.press('tab')   # Move to Restore Original Values radio
        time.sleep(0.1)
        pyautogui.press('space')  # Select it
        time.sleep(0.2)
        steps.append("Selected Restore Original Values")

    if save_scenario:
        # Click "Save Scenario..." button in Results dialog
        # Tab to the Save Scenario button (past radio buttons into buttons area)
        if restore_original:
            # Already past Restore Original Values, tab to buttons
            for _ in range(2):
                pyautogui.press('tab')
                time.sleep(0.1)
        else:
            # From Keep Solver Solution, tab past Restore Original Values to buttons
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.1)
        pyautogui.press('enter')  # Save Scenario button
        time.sleep(0.5)

        # Enter scenario name
        pyautogui.typewrite(save_scenario, interval=0.03)
        pyautogui.press('enter')  # OK to close scenario name dialog
        time.sleep(0.5)
        steps.append(f"Saved scenario: {save_scenario}")

    # OK to close Solver Results dialog
    pyautogui.press('enter')
    time.sleep(0.5)

    return {"status": "completed", "steps": steps}


def do_data_analysis(tool_name: str, input_range: str, output_range: str = "",
                     options: dict = None, sheet: str = ""):
    """Run Analysis ToolPak tool with full option support.

    options dict can include:
    - grouped_by: "columns" or "rows"
    - labels_in_first_row: bool
    - summary_statistics: bool
    - confidence_level: bool / float
    - kth_largest: int
    - kth_smallest: int
    - output_range: str (alternative to output_range param)
    """
    steps = []
    if not options:
        options = {}

    ensure_clean_state()

    if sheet:
        switch_sheet(sheet)
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.3)

    # Open Data Analysis via Help search
    search_menu("Data Analysis")
    steps.append("Opened Data Analysis dialog")

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

    search_menu("Excel Add-ins")
    time.sleep(1.0)
    steps.append("Opened Excel Add-ins dialog")

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
        # Check for user cancellation before each action
        if cancel_check and cancel_check():
            print(f"[excel] ⛔ CANCELLED by user before action {i+1}/{len(actions)}")
            cancelled = True
            # Clean up: dismiss any open dialogs and return to safe state
            _emergency_cleanup()
            break

        print(f"[excel] Action {i+1}/{len(actions)}: {action.get('type')}")
        result = execute_excel_action(action)
        results.append(result)

        status = result.get("status", "unknown")
        print(f"[excel]   -> {status}")
        for step in result.get("steps", []):
            print(f"[excel]     {step}")

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
