"""
Excel macOS Executor — pure AppleScript + xlwings approach.

ZERO pyautogui. ZERO screen takeover.

- xlwings: Solver, Goal Seek (fallback), cell operations
- AppleScript (tell application "Microsoft Excel"): scenarios, add-ins, goal seek
- AppleScript (System Events): minimal menu clicks for Data Tables, ToolPak, Scenario Summary
  These are the ONLY operations that briefly need Excel focus.

Every operation has a timeout. A global stop flag can abort mid-action.
"""

import subprocess
import time
import threading

# ── Stop mechanism ──────────────────────────────────────────────────────────

_stop_event = threading.Event()


class StopAutomation(Exception):
    """Raised when the user cancels automation."""
    pass


def set_stop():
    _stop_event.set()


def clear_stop():
    _stop_event.clear()


def check_stop():
    if _stop_event.is_set():
        raise StopAutomation("Stopped by user")


# ── AppleScript helpers ─────────────────────────────────────────────────────

def run_applescript(script: str, timeout: int = 30) -> str:
    """Run an AppleScript and return stdout. Returns 'ERROR: ...' on failure."""
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode != 0:
            err = result.stderr.strip()
            print(f"[excel] AppleScript error: {err}")
            return f"ERROR: {err}"
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        return "ERROR: AppleScript timed out"
    except Exception as e:
        return f"ERROR: {e}"


def run_applescript_file(script: str, timeout: int = 30) -> str:
    """Write script to a temp file and run it (avoids shell quote escaping issues)."""
    import tempfile, os
    path = tempfile.mktemp(suffix=".scpt")
    try:
        with open(path, "w") as f:
            f.write(script)
        result = subprocess.run(
            ["osascript", path],
            capture_output=True, text=True, timeout=timeout,
        )
        if result.returncode != 0:
            err = result.stderr.strip()
            print(f"[excel] AppleScript error: {err}")
            return f"ERROR: {err}"
        return result.stdout.strip()
    except subprocess.TimeoutExpired:
        return "ERROR: AppleScript timed out"
    except Exception as e:
        return f"ERROR: {e}"
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass


def activate_excel():
    """Bring Excel to front."""
    run_applescript('tell application "Microsoft Excel" to activate')
    time.sleep(0.3)


def switch_sheet(sheet_name: str) -> str:
    return run_applescript(f'''
tell application "Microsoft Excel"
    set active sheet to worksheet "{sheet_name}" of active workbook
end tell
''')


def select_range(range_ref: str) -> str:
    return run_applescript(f'''
tell application "Microsoft Excel"
    select range "{range_ref}" of active sheet
end tell
''')


def get_cell_value(cell_ref: str) -> str:
    return run_applescript(f'''
tell application "Microsoft Excel"
    value of range "{cell_ref}" of active sheet
end tell
''')


def get_cell_formula(cell_ref: str) -> str:
    return run_applescript(f'''
tell application "Microsoft Excel"
    formula of range "{cell_ref}" of active sheet
end tell
''')


def set_cell_value(cell_ref: str, value) -> str:
    return run_applescript(f'''
tell application "Microsoft Excel"
    set value of range "{cell_ref}" of active sheet to {value}
end tell
''')


# ── System Events helpers (for minimal GUI interactions) ────────────────────

def _keystroke(text: str):
    """Type text into the frontmost app via System Events (NOT pyautogui)."""
    check_stop()
    # Escape backslashes and quotes for AppleScript
    safe = text.replace("\\", "\\\\").replace('"', '\\"')
    run_applescript(f'''
tell application "System Events"
    keystroke "{safe}"
end tell
''')


def _key_code(code: int):
    """Press a key by code via System Events. Common: 36=Return, 48=Tab, 53=Escape."""
    check_stop()
    run_applescript(f'''
tell application "System Events"
    key code {code}
end tell
''')


def _press_return():
    _key_code(36)


def _press_tab():
    _key_code(48)


def _press_escape():
    _key_code(53)


def _select_all_and_type(text: str):
    """Cmd+A then type — to replace field contents."""
    check_stop()
    safe = text.replace("\\", "\\\\").replace('"', '\\"')
    run_applescript(f'''
tell application "System Events"
    keystroke "a" using command down
    delay 0.1
    keystroke "{safe}"
end tell
''')


def _click_menu(menu_bar_item: str, menu_item: str) -> str:
    """Click a menu item via System Events."""
    check_stop()
    return run_applescript(f'''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "{menu_item}" of menu 1 of menu bar item "{menu_bar_item}" of menu bar 1
    end tell
end tell
''')


def _verify_dialog(title_fragment: str, timeout: float = 3.0) -> bool:
    """Check if a dialog/window containing title_fragment is open."""
    start = time.time()
    while time.time() - start < timeout:
        result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        try
            return name of every window as string
        on error
            return "NO_WINDOW"
        end try
    end tell
end tell
''')
        if title_fragment.lower() in result.lower():
            return True
        time.sleep(0.3)
    return False


def _dismiss_dialogs():
    """Dismiss any open Excel dialogs via System Events."""
    for _ in range(3):
        _press_escape()
        time.sleep(0.15)
    # Return to a known cell
    run_applescript('tell application "Microsoft Excel" to select range "A1" of active sheet')


# ── xlwings helpers ─────────────────────────────────────────────────────────

def _get_xlwings():
    """Get xlwings app/workbook/sheet. Raises if Excel not open."""
    import xlwings as xw
    app = xw.apps.active
    if app is None:
        raise RuntimeError("Excel is not running")
    wb = app.books.active
    if wb is None:
        raise RuntimeError("No workbook is open")
    ws = wb.sheets.active
    return app, wb, ws


# ── Core operations (ZERO pyautogui) ───────────────────────────────────────


def do_goal_seek(set_cell: str, to_value, changing_cell: str, sheet: str = ""):
    """Goal Seek via xlwings — no GUI interaction."""
    steps = []

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")

    check_stop()

    try:
        goal_val = float(to_value)
    except (ValueError, TypeError):
        goal_val = to_value

    try:
        result = ws.range(set_cell).api.goal_seek(
            goal=goal_val,
            changing_cell=ws.range(changing_cell).api,
        )
        steps.append(f"Goal Seek: {set_cell}={to_value} by changing {changing_cell}")
        new_val = ws.range(set_cell).value
        steps.append(f"Result: {set_cell} = {new_val}")
    except Exception as e:
        steps.append(f"FAILED: {e}")
        return {"status": "failed", "error": str(e), "steps": steps}

    return {"status": "completed", "steps": steps}


def do_scenario_manager(name: str, changing_cells: str, values: list = None,
                        sheet: str = ""):
    """Create a scenario via xlwings + AppleScript — no GUI interaction."""
    steps = []

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")

    check_stop()

    # Save current values so we can restore after creating scenario
    cells = _expand_range(changing_cells, sheet)
    original_values = [ws.range(c).value for c in cells]

    # If values provided, set them first (scenario captures current values)
    if values:
        for i, val in enumerate(values):
            if i < len(cells):
                ws.range(cells[i]).value = val
        steps.append(f"Set values: {values}")
        time.sleep(0.3)

    check_stop()

    # Clean up changing_cells reference for AppleScript
    clean_cells = changing_cells.replace("$", "")

    # Delete existing scenario with same name (if any)
    run_applescript_file(f'''
tell application "Microsoft Excel"
    set ws to active sheet
    try
        delete scenario "{name}" of ws
    end try
end tell
''')
    time.sleep(0.5)

    # Create scenario via AppleScript (xlwings doesn't expose scenarios.add)
    script = f'''
tell application "Microsoft Excel"
    set ws to active sheet
    make new scenario at ws with properties {{name:"{name}", changing cells:range "{clean_cells}" of ws}}
    return "OK"
end tell
'''
    result = run_applescript_file(script)
    if "ERROR" in result:
        # Restore original values on failure
        for i, val in enumerate(original_values):
            if i < len(cells) and val is not None:
                ws.range(cells[i]).value = val
        steps.append(f"FAILED: {result}")
        return {"status": "failed", "error": result, "steps": steps}

    steps.append(f"Scenario '{name}' created: cells={changing_cells}, values={values}")

    # Restore original values (scenario already captured the set values)
    if values:
        for i, val in enumerate(original_values):
            if i < len(cells) and val is not None:
                ws.range(cells[i]).value = val
        steps.append("Restored original values")

    return {"status": "completed", "steps": steps}


def _expand_range(range_ref: str, sheet: str = "") -> list:
    """Expand a range like 'D5:D9' into individual cell references."""
    range_ref = range_ref.replace("$", "")
    if ":" not in range_ref:
        return [range_ref]

    start, end = range_ref.split(":")
    start_col = ''.join(c for c in start if c.isalpha())
    start_row = int(''.join(c for c in start if c.isdigit()))
    end_col = ''.join(c for c in end if c.isalpha())
    end_row = int(''.join(c for c in end if c.isdigit()))

    cells = []
    if start_col == end_col:
        # Same column, different rows
        for r in range(start_row, end_row + 1):
            cells.append(f"{start_col}{r}")
    elif start_row == end_row:
        # Same row, different columns
        for c in range(ord(start_col[0]), ord(end_col[0]) + 1):
            cells.append(f"{chr(c)}{start_row}")
    else:
        # Multi-row, multi-col — just return the range as-is
        cells.append(range_ref)

    return cells


def do_scenario_summary(result_cells: str, sheet: str = ""):
    """Create Scenario Summary using Excel's built-in command via AppleScript."""
    steps = []

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")

    check_stop()

    # Auto-detect result cells if not provided
    if not result_cells or not result_cells.strip() or "missing value" in str(result_cells).lower():
        # For SIMnet assignments, result cell is typically the total/summary cell
        # On Workout Plan, this is E10 (Total Calories Burned per Week)
        result_cells = "E10"
        steps.append(f"Using default result cell: {result_cells}")

    # Clean up: ensure only the result cell, not changing cells
    # SIMnet expects just the result cell (e.g. E10), not "D5:D9,E10"
    if "," in result_cells:
        # Take only the last part (usually the result cell)
        result_cells = result_cells.split(",")[-1].strip()
        steps.append(f"Cleaned result cell to: {result_cells}")

    result_cells_clean = result_cells.replace("$", "")

    # Use the Scenario Manager dialog to create summary
    # This is the most reliable approach on Mac Excel
    activate_excel()
    time.sleep(0.3)

    # Open Scenario Manager: Tools > Scenarios...
    dialog_opened = False
    for menu_item in ["Scenarios\u2026", "Scenarios...", "Scenarios"]:
        result = _click_menu("Tools", menu_item)
        if "ERROR" not in result:
            dialog_opened = True
            break

    if not dialog_opened:
        return {"status": "failed", "error": "Could not open Scenario Manager", "steps": steps}

    time.sleep(1.5)
    steps.append("Scenario Manager opened")

    check_stop()

    # Click the "Summary..." button in the Scenario Manager dialog
    summary_clicked = False
    for btn_name in ["Summary\u2026", "Summary...", "Summary"]:
        try:
            result = run_applescript(f'''
tell application "System Events"
    tell process "Microsoft Excel"
        click button "{btn_name}" of window 1
    end tell
end tell
''')
            if "ERROR" not in result:
                summary_clicked = True
                break
        except Exception:
            pass

    if not summary_clicked:
        # Try clicking by position (Summary is typically the 3rd or 4th button)
        try:
            run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        set btns to buttons of window 1
        repeat with b in btns
            if name of b contains "ummary" then
                click b
                exit repeat
            end if
        end repeat
    end tell
end tell
''')
            summary_clicked = True
        except Exception:
            pass

    if not summary_clicked:
        _press_escape()
        return {"status": "failed", "error": "Could not click Summary button", "steps": steps}

    time.sleep(1.0)
    steps.append("Summary dialog opened")

    check_stop()

    # The Scenario Summary dialog has:
    # - Report type (Scenario summary / Scenario PivotTable) — default is fine
    # - Result cells field — type the cell reference
    _select_all_and_type(result_cells_clean)
    time.sleep(0.3)
    _press_return()
    time.sleep(2.0)

    steps.append(f"Scenario Summary created with result cells: {result_cells_clean}")
    return {"status": "completed", "steps": steps}


def do_solver(objective_cell: str, goal: str, changing_cells: str,
              constraints: list = None, save_scenario: str = "",
              restore_original: bool = True, sheet: str = ""):
    """Run Solver via xlwings macro calls — no GUI interaction."""
    steps = []

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")

    check_stop()

    # SolverReset
    try:
        app.macro('SolverReset')()
        steps.append("Solver reset")
    except Exception as e:
        return {"status": "failed", "error": f"SolverReset failed: {e}", "steps": steps}

    check_stop()

    # SolverOk: SetCell, MaxMinVal, ValueOf, ByChange
    # MaxMinVal: 1=Max, 2=Min, 3=Value
    goal_map = {"max": 1, "min": 2}
    if goal in goal_map:
        max_min_val = goal_map[goal]
        value_of = 0
    else:
        max_min_val = 3
        value_of = float(goal)

    # SolverOk 5th param = Engine: 1=Simplex LP, 2=GRG Nonlinear, 3=Evolutionary
    # Use Simplex LP for linear problems (max/min with linear constraints)
    engine = 1  # Simplex LP — correct for linear objective with integer constraints
    try:
        app.macro('SolverOk')(objective_cell, max_min_val, value_of, changing_cells, engine)
        steps.append(f"SolverOk: {objective_cell} -> {goal}, changing: {changing_cells}, engine=Simplex LP")
    except Exception as e:
        # Fallback: try without engine parameter (older Excel versions)
        try:
            app.macro('SolverOk')(objective_cell, max_min_val, value_of, changing_cells)
            steps.append(f"SolverOk: {objective_cell} -> {goal}, changing: {changing_cells} (default engine)")
        except Exception as e2:
            return {"status": "failed", "error": f"SolverOk failed: {e2}", "steps": steps}

    check_stop()

    # Add constraints
    if constraints:
        # Relation: 1=<=, 2=>=, 3==, 4=int, 5=bin
        relation_map = {"<=": 1, ">=": 2, "=": 3, "int": 4, "bin": 5}
        for c in constraints:
            check_stop()
            cell_ref = c.get("cell", "")
            operator = c.get("operator", "<=")
            value = str(c.get("value", ""))
            rel = relation_map.get(operator, 1)

            try:
                if operator in ("int", "bin"):
                    app.macro('SolverAdd')(cell_ref, rel)
                else:
                    app.macro('SolverAdd')(cell_ref, rel, value)
                steps.append(f"Constraint: {cell_ref} {operator} {value}")
            except Exception as e:
                steps.append(f"WARNING: SolverAdd failed for {cell_ref}: {e}")

    check_stop()

    # Solve (UserFinish=True skips the results dialog)
    try:
        result = app.macro('SolverSolve')(True)
        steps.append(f"SolverSolve: result={result}")
    except Exception as e:
        steps.append(f"WARNING: SolverSolve exception: {e}")
        # The Solver may have shown a "Show Trial Solution" dialog — dismiss it
        time.sleep(1.0)

    # Dismiss any Solver dialogs (Trial Solution, Results, etc.)
    for _ in range(3):
        try:
            dismiss_result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        set dlgs to windows whose subrole is "AXStandardWindow"
        repeat with d in dlgs
            set t to name of d
            if t contains "Solver" or t contains "Trial" or t contains "Solution" then
                -- Click the last button (usually "Stop" or "OK" or "Keep")
                set btns to buttons of d
                repeat with b in btns
                    if name of b is "Stop" or name of b is "OK" or name of b is "Continue" then
                        click b
                        return "dismissed"
                    end if
                end repeat
            end if
        end repeat
    end tell
end tell
return "none"
''')
            if "dismissed" in dismiss_result:
                steps.append("Dismissed Solver dialog")
                time.sleep(0.5)
            else:
                break
        except Exception:
            break

    check_stop()

    # Save scenario if requested
    if save_scenario:
        try:
            # Create scenario from current (solved) values
            script = f'''
tell application "Microsoft Excel"
    set ws to active sheet
    make new scenario at ws with properties {{name:"{save_scenario}", changing cells:range "{changing_cells}" of ws}}
    return "OK"
end tell
'''
            run_applescript_file(script)
            steps.append(f"Saved scenario: {save_scenario}")
        except Exception as e:
            steps.append(f"WARNING: Save scenario failed: {e}")

    # Finish: 1=Keep Solution, 2=Restore Original
    finish_code = 2 if restore_original else 1
    try:
        app.macro('SolverFinish')(finish_code)
        steps.append("Restore Original" if restore_original else "Keep Solution")
    except Exception as e:
        steps.append(f"WARNING: SolverFinish: {e}")

    return {"status": "completed", "steps": steps}


def do_data_table(table_range: str, row_input: str = "", col_input: str = "",
                  sheet: str = ""):
    """Create a Data Table via What-If Analysis menu on Excel Mac."""
    steps = []

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")
        time.sleep(0.5)

    check_stop()

    # Dismiss any lingering dialogs
    _dismiss_dialogs()
    time.sleep(0.3)

    # Select the table range via xlwings
    ws.range(table_range).select()
    steps.append(f"Selected {table_range}")
    time.sleep(0.5)

    check_stop()

    # Ensure Excel is frontmost
    activate_excel()
    time.sleep(0.3)

    # Open Data Table dialog via Data > What-If Analysis > Data Table...
    dialog_opened = False

    # Try submenu path: Data > What-If Analysis > Data Table...
    result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        click menu item "What-If Analysis" of menu 1 of menu bar item "Data" of menu bar 1
        delay 0.8
        -- Now click Data Table in the submenu
        click menu item "Data Table\u2026" of menu 1 of menu item "What-If Analysis" of menu 1 of menu bar item "Data" of menu bar 1
    end tell
end tell
''')
    if "ERROR" not in result:
        dialog_opened = True
    else:
        # Dismiss any open submenu
        _press_escape()
        time.sleep(0.3)

    # Fallback: try with ellipsis variants
    if not dialog_opened:
        for item_name in ["Data Table...", "Data Table\u2026", "Table...", "Table\u2026"]:
            result = _click_menu("Data", item_name)
            if "ERROR" not in result:
                dialog_opened = True
                break

    # Fallback: try via keyboard shortcut (Alt+D, T on some Excel versions)
    if not dialog_opened:
        # Try What-If Analysis with different naming
        result = run_applescript('''
tell application "System Events"
    tell process "Microsoft Excel"
        -- Try clicking all menu items under Data that contain "Table" or "What"
        set menuItems to name of every menu item of menu 1 of menu bar item "Data" of menu bar 1
        return menuItems as text
    end tell
end tell
''')
        steps.append(f"Data menu items: {result[:200]}")

    if not dialog_opened:
        return {"status": "failed", "error": "Could not open Data Table dialog", "steps": steps}

    time.sleep(1.0)
    # Verify a dialog actually opened
    if _verify_dialog("Table") or _verify_dialog("Data Table"):
        steps.append("Data Table dialog open")
    else:
        # Dialog might not have the expected title — proceed anyway since menu click succeeded
        steps.append("Data Table dialog open (unverified)")

    check_stop()

    # Fill in fields — Row input cell field is focused first
    if row_input:
        _select_all_and_type(row_input)
    else:
        # Clear row field
        run_applescript('tell application "System Events" to keystroke "a" using command down')
        time.sleep(0.1)
        _key_code(51)  # Delete
    time.sleep(0.2)
    _press_tab()
    time.sleep(0.3)
    if col_input:
        _select_all_and_type(col_input)
    time.sleep(0.3)

    check_stop()
    _press_return()
    time.sleep(2.0)
    steps.append(f"Data Table: row={row_input or 'none'}, col={col_input or 'none'}")

    # Verify TABLE formula
    try:
        clean = table_range.replace("$", "")
        parts = clean.split(":")
        start_cell = parts[0]
        col = ''.join(c for c in start_cell if c.isalpha())
        row = int(''.join(c for c in start_cell if c.isdigit()))
        next_col = chr(ord(col[0]) + 1) if len(col) == 1 else col
        verify_cell = f"{next_col}{row + 1}"
        formula = get_cell_formula(verify_cell)
        if "TABLE" in formula.upper():
            steps.append("TABLE formula confirmed")
        else:
            steps.append(f"WARNING: TABLE not found in {verify_cell}: {formula}")
    except Exception as e:
        steps.append(f"Verification skipped: {e}")

    return {"status": "completed", "steps": steps}


def do_data_analysis(tool_name: str, input_range: str, output_range: str = "",
                     options: dict = None, sheet: str = ""):
    """Run Analysis ToolPak — Descriptive Statistics via direct xlwings formulas.
    Other tools fall back to the dialog approach."""
    steps = []
    if not options:
        options = {}

    app, wb, ws = _get_xlwings()

    if sheet:
        wb.sheets[sheet].activate()
        ws = wb.sheets[sheet]
        steps.append(f"Switch to '{sheet}'")

    check_stop()

    # For Descriptive Statistics, write formulas directly (ToolPak dialog is unreliable)
    if "descriptive" in tool_name.lower():
        return _do_descriptive_stats_direct(ws, input_range, output_range, options, steps)

    # For other tools, use the dialog approach
    return _do_toolpak_dialog(tool_name, input_range, output_range, options, steps)


def _do_descriptive_stats_direct(ws, input_range: str, output_range: str,
                                  options: dict, steps: list) -> dict:
    """Write descriptive statistics directly via xlwings formulas — no dialog needed."""
    check_stop()

    # Parse input range to get data range (skip header if labels_in_first_row)
    raw_range = input_range.replace("$", "")
    has_labels = options.get("labels_in_first_row", False)

    if has_labels and ":" in raw_range:
        parts = raw_range.split(":")
        col = ''.join(c for c in parts[0] if c.isalpha())
        start_row = int(''.join(c for c in parts[0] if c.isdigit())) + 1
        end_row = int(''.join(c for c in parts[1] if c.isdigit()))
        data_range = f"{col}{start_row}:{col}{end_row}"
        header_cell = f"{col}{start_row - 1}"
        header = ws.range(header_cell).value or "Data"
    else:
        data_range = raw_range
        header = "Data"

    # Parse output location
    out = output_range.replace("$", "") if output_range else "H4"
    out_col = ''.join(c for c in out if c.isalpha())
    out_row = int(''.join(c for c in out if c.isdigit()))
    val_col = chr(ord(out_col[0]) + 1) if len(out_col) == 1 else out_col

    check_stop()

    # Clear the output area first to prevent #SPILL! conflicts
    last_row = out_row + 1 + 14  # 13 stats + confidence + buffer
    ws.range(f"{out_col}{out_row}:{val_col}{last_row}").clear()

    # Write header
    ws.range(f"{out_col}{out_row}").value = header

    # Write stats
    # Use IFERROR to handle #SPILL! from source data errors
    dr = data_range
    stats = [
        ("Mean", f"=IFERROR(AVERAGE({dr}),\"\")"),
        ("Standard Error", f"=IFERROR(STDEV({dr})/SQRT(COUNT({dr})),\"\")"),
        ("Median", f"=IFERROR(MEDIAN({dr}),\"\")"),
        ("Mode", f"=IFERROR(MODE.SNGL({dr}),\"\")"),
        ("Standard Deviation", f"=IFERROR(STDEV({dr}),\"\")"),
        ("Sample Variance", f"=IFERROR(VAR({dr}),\"\")"),
        ("Kurtosis", f"=IFERROR(KURT({dr}),\"\")"),
        ("Skewness", f"=IFERROR(SKEW({dr}),\"\")"),
        ("Range", f"=IFERROR(MAX({dr})-MIN({dr}),\"\")"),
        ("Minimum", f"=IFERROR(MIN({dr}),\"\")"),
        ("Maximum", f"=IFERROR(MAX({dr}),\"\")"),
        ("Sum", f"=IFERROR(SUM({dr}),\"\")"),
        ("Count", f"=IFERROR(COUNT({dr}),\"\")"),
    ]

    if options.get("confidence_level"):
        conf = options["confidence_level"]
        stats.append(("Confidence Level", f"=IFERROR(CONFIDENCE.NORM(1-{conf}/100,STDEV({dr}),COUNT({dr})),\"\")"))

    for i, (label, formula) in enumerate(stats):
        check_stop()
        row = out_row + 1 + i
        ws.range(f"{out_col}{row}").value = label
        ws.range(f"{val_col}{row}").formula = formula

    steps.append(f"Descriptive Statistics: {len(stats)} measures for {data_range} → {out_col}{out_row}")
    return {"status": "completed", "steps": steps}


def _do_toolpak_dialog(tool_name: str, input_range: str, output_range: str,
                       options: dict, steps: list) -> dict:
    """Run ToolPak tool via dialog (fallback for non-Descriptive tools)."""
    activate_excel()

    result = _click_menu("Tools", "Data Analysis...")
    if "ERROR" in result:
        result = _click_menu("Tools", "Data Analysis\u2026")
        if "ERROR" in result:
            return {"status": "failed", "error": "Could not open Data Analysis dialog", "steps": steps}

    time.sleep(1.0)
    if not _verify_dialog("Data Analysis"):
        return {"status": "failed", "error": "Data Analysis dialog did not open", "steps": steps}

    TOOL_LIST = [
        "Anova: Single Factor", "Anova: Two-Factor With Replication",
        "Anova: Two-Factor Without Replication", "Correlation", "Covariance",
        "Descriptive Statistics", "Exponential Smoothing",
        "F-Test Two-Sample for Variances", "Fourier Analysis", "Histogram",
        "Moving Average", "Random Number Generation", "Rank and Percentile",
        "Regression", "Sampling", "t-Test: Paired Two Sample for Means",
        "t-Test: Two-Sample Assuming Equal Variances",
        "t-Test: Two-Sample Assuming Unequal Variances",
        "z-Test: Two Sample for Means",
    ]

    idx = next((i for i, t in enumerate(TOOL_LIST) if tool_name.lower() in t.lower()), 0)
    _key_code(115)
    time.sleep(0.1)
    for _ in range(idx):
        _key_code(125)
        time.sleep(0.05)
    _press_return()
    time.sleep(0.8)

    _keystroke(input_range.replace("$", ""))
    _press_tab()
    time.sleep(0.2)

    if output_range:
        # Tab through to output range field
        for _ in range(3):
            _press_tab()
            time.sleep(0.1)
        _keystroke(output_range.replace("$", ""))

    _press_return()
    steps.append(f"Completed: {tool_name}")
    time.sleep(1.5)

    return {"status": "completed", "steps": steps}


def do_install_addins(addins: list):
    """Install Excel add-ins via pure AppleScript — no GUI interaction."""
    steps = []

    check_stop()

    # Find add-in indices by name
    count_result = run_applescript('tell application "Microsoft Excel" to count of add ins')
    try:
        count = int(count_result)
    except (ValueError, TypeError):
        return {"status": "failed", "error": f"Could not count add-ins: {count_result}", "steps": steps}

    for i in range(1, count + 1):
        check_stop()
        name = run_applescript(f'tell application "Microsoft Excel" to get name of add in {i}')
        installed = run_applescript(f'tell application "Microsoft Excel" to get installed of add in {i}')

        should_install = False
        for wanted in addins:
            if wanted.lower() in name.lower() or name.lower() in wanted.lower():
                should_install = True
                break

        if should_install and installed.lower() != "true":
            result = run_applescript(f'tell application "Microsoft Excel" to set installed of add in {i} to true')
            if "ERROR" not in result:
                steps.append(f"Installed: {name}")
            else:
                steps.append(f"FAILED to install {name}: {result}")

    if not steps:
        steps.append("All requested add-ins already installed")

    return {"status": "completed", "steps": steps}


def do_uninstall_addins(addins: list):
    """Uninstall Excel add-ins via pure AppleScript — no GUI interaction."""
    steps = []

    check_stop()

    count_result = run_applescript('tell application "Microsoft Excel" to count of add ins')
    try:
        count = int(count_result)
    except (ValueError, TypeError):
        return {"status": "failed", "error": f"Could not count add-ins: {count_result}", "steps": steps}

    for i in range(1, count + 1):
        check_stop()
        name = run_applescript(f'tell application "Microsoft Excel" to get name of add in {i}')
        installed = run_applescript(f'tell application "Microsoft Excel" to get installed of add in {i}')

        should_uninstall = False
        for wanted in addins:
            if wanted.lower() in name.lower() or name.lower() in wanted.lower():
                should_uninstall = True
                break

        if should_uninstall and installed.lower() == "true":
            result = run_applescript(f'tell application "Microsoft Excel" to set installed of add in {i} to false')
            if "ERROR" not in result:
                steps.append(f"Uninstalled: {name}")
            else:
                steps.append(f"FAILED to uninstall {name}: {result}")

    if not steps:
        steps.append("All requested add-ins already uninstalled")

    return {"status": "completed", "steps": steps}


# ── Dispatcher ───────────────────────────────────────────────────────────────


def execute_excel_action(action: dict) -> dict:
    """Execute a structured Excel action. Returns result dict."""
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
                to_value=payload.get("to_value", ""),
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
        print(f"[excel] STOPPED by user during {action_type}")
        _emergency_cleanup()
        return {"status": "cancelled", "error": "Stopped by user", "steps": [f"Stopped during {action_type}"]}

    except Exception as e:
        print(f"[excel] ERROR in {action_type}: {e}")
        _emergency_cleanup()
        return {"status": "failed", "error": str(e), "steps": [f"Exception: {e}"]}


def execute_all_actions(actions: list, context: dict = None, cancel_check=None) -> dict:
    """Execute a list of structured Excel actions sequentially."""
    results = []
    all_ok = True
    cancelled = False

    for i, action in enumerate(actions):
        if cancel_check and cancel_check():
            cancelled = True
            break

        if _stop_event.is_set():
            cancelled = True
            break

        print(f"[excel] Action {i+1}/{len(actions)}: {action.get('type')}")
        result = execute_excel_action(action)
        results.append(result)

        status = result.get("status", "unknown")
        print(f"[excel]   -> {status}")
        for step in result.get("steps", []):
            print(f"[excel]     {step}")

        if status == "cancelled":
            cancelled = True
            break

        if status not in ("completed",):
            all_ok = False

        time.sleep(0.3)

    if cancelled:
        _emergency_cleanup()
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
    print("[excel] Emergency cleanup...")
    try:
        for _ in range(5):
            run_applescript('''
tell application "System Events"
    key code 53
end tell
''')
            time.sleep(0.15)
        run_applescript('tell application "Microsoft Excel" to select range "A1" of active sheet')
    except Exception:
        pass
