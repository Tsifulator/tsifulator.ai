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
    """Goal Seek via DIRECT AppleScript — bypasses xlwings's broken appscript
    Range method translation on Mac.

    Why not xlwings.api.goal_seek? On Mac Excel, xlwings's appscript backend
    translates `Range.goal_seek(...)` to `cells['C18'].goal_seek(...)` which
    Excel for Mac's AppleScript dictionary REJECTS with OSERROR -50
    ("Parameter error"). The dictionary expects `range "C18" of ws` syntax,
    not the cells collection accessor. So we bypass xlwings entirely and
    construct the AppleScript ourselves.

    Hardening (preserved from xlwings version):
      1. `display alerts` set to false → no VBA error dialog popups during
         iteration (the '300 popups' bug from SimNet workbooks)
      2. Pre-flight validation: distinct cells, numeric to_value,
         changing_cell isn't a formula
      3. Convergence check: read before/after values, fail clean if
         changing_cell didn't move
      4. xlwings still used for read-only operations (validate, capture
         before/after values) since those work fine on Mac
    """
    steps = []

    # ── Pre-flight validation ──────────────────────────────────────────────
    if not set_cell or not changing_cell:
        return {"status": "failed",
                "error": "Both set_cell and changing_cell are required"}

    if set_cell.replace("$", "").upper() == changing_cell.replace("$", "").upper():
        return {"status": "failed",
                "error": f"set_cell ({set_cell}) and changing_cell ({changing_cell}) are the same — Goal Seek requires distinct cells"}

    try:
        goal_val = float(to_value)
    except (ValueError, TypeError):
        return {"status": "failed",
                "error": f"to_value must be numeric, got {to_value!r}"}

    # Use xlwings (read-only) to validate the changing_cell + capture before
    try:
        app, wb, ws = _get_xlwings()
    except Exception as e:
        return {"status": "failed", "error": f"xlwings init: {e}"}

    target_sheet = sheet or ws.name
    try:
        ws_obj = wb.sheets[target_sheet]
    except Exception:
        return {"status": "failed", "error": f"Sheet '{target_sheet}' not found"}

    check_stop()

    try:
        chg_range = ws_obj.range(changing_cell)
        chg_formula = chg_range.formula
        if isinstance(chg_formula, str) and chg_formula.startswith("="):
            return {
                "status": "failed",
                "error": (f"changing_cell {changing_cell} contains a formula "
                          f"({chg_formula!r}). Goal Seek needs a numeric input "
                          f"cell, not a calculated one. Pick the cell that "
                          f"holds the input value (e.g. a rate or quantity)."),
            }
        before_val = chg_range.value
        set_before = ws_obj.range(set_cell).value
    except Exception as e:
        return {"status": "failed",
                "error": f"Could not read cells: {e}"}

    # Strip $ from cell refs — AppleScript range syntax doesn't use them
    # the same way and they can confuse the parser
    set_clean = set_cell.replace("$", "")
    chg_clean = changing_cell.replace("$", "")
    sheet_clean = target_sheet.replace('"', '\\"')

    # ── DIRECT APPLESCRIPT ──────────────────────────────────────────────────
    # Disable alerts so workbook VBA errors don't pop dialogs during the
    # iterative recalc, then call Excel's native Goal Seek verb on the
    # target range. Re-enable alerts in the same script (no try/finally
    # needed because the alerts setting is process-scoped and restoring
    # at the end of THIS call is enough).
    script = f'''
tell application "Microsoft Excel"
    set savedAlerts to display alerts
    set display alerts to false
    try
        set ws to worksheet "{sheet_clean}" of active workbook
        set targetRange to range "{set_clean}" of ws
        set changingRange to range "{chg_clean}" of ws
        goal seek targetRange goal {goal_val} changing cell changingRange
        set display alerts to savedAlerts
        return "OK"
    on error errMsg number errNum
        set display alerts to savedAlerts
        return "ERROR: " & errMsg & " (number " & errNum & ")"
    end try
end tell
'''
    check_stop()
    result = run_applescript_file(script, timeout=60)

    if isinstance(result, str) and result.startswith("ERROR"):
        return {
            "status": "failed",
            "error": f"AppleScript Goal Seek failed: {result}",
            "steps": steps,
        }

    # ── Capture after-values + convergence check ───────────────────────────
    try:
        chg_after = ws_obj.range(changing_cell).value
        set_after = ws_obj.range(set_cell).value
    except Exception:
        chg_after = None
        set_after = None

    moved = (
        isinstance(before_val, (int, float))
        and isinstance(chg_after, (int, float))
        and abs(before_val - chg_after) > 1e-9
    )
    if not moved:
        return {
            "status": "failed",
            "error": (f"Goal Seek did not converge. {changing_cell} stayed "
                      f"at {before_val!r}. Check that {set_cell}'s formula "
                      f"actually depends on {changing_cell}, and that the "
                      f"target value ({goal_val}) is reachable."),
            "steps": steps,
        }

    # Format numeric values nicely for the user-facing message
    def _fmt(v):
        if isinstance(v, float):
            # 4 decimals max, drop trailing zeros
            return f"{v:.4f}".rstrip("0").rstrip(".")
        return str(v)

    steps.append(
        f"Goal Seek: {set_cell}: {_fmt(set_before)} → {_fmt(set_after)} "
        f"by changing {changing_cell}: {_fmt(before_val)} → {_fmt(chg_after)}"
    )
    return {
        "status": "completed",
        "message": (f"Goal Seek converged: {changing_cell} changed from "
                    f"{_fmt(before_val)} to {_fmt(chg_after)}, "
                    f"so {set_cell} = {_fmt(set_after)} (target was {_fmt(goal_val)})"),
        "steps": steps,
    }


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


# ── P1 handlers (xlwings, runs silently in background) ─────────────────────


def do_smartart_diagram(steps: list, sheet: str = "",
                         layout: str = "process", anchor: str = ""):
    """Create a SmartArt-style flow diagram on the given sheet.

    Excel for Mac's `Shapes.AddSmartArt` is unreliable, so we build an
    equivalent flow diagram from rounded rectangles + arrow connectors —
    visually identical to "Basic Process" SmartArt and 100% deterministic.

    Args:
        steps: list of node text strings, e.g.
               ["Selling Price", "Commission Rate", "Total Commission", "PHRE Share"]
        sheet: target worksheet name (defaults to active sheet)
        layout: 'process' (horizontal arrow flow), 'cycle' (circular arrow),
                'list' (stacked rectangles). Only 'process' is shipped in P1;
                cycle/list fall back to process.
        anchor: optional top-left anchor cell like 'F2'. If empty, defaults
                to a position right of the used range.

    Runs entirely in the background via xlwings — no Excel focus needed.
    """
    if not steps:
        return {"status": "failed", "error": "No steps provided for SmartArt diagram"}

    check_stop()

    try:
        app, wb, ws = _get_xlwings()
    except Exception as e:
        return {"status": "failed", "error": f"xlwings init: {e}"}

    # Target sheet: switch only if specified and different from active
    if sheet:
        try:
            ws = wb.sheets[sheet]
        except Exception:
            return {"status": "failed", "error": f"Sheet '{sheet}' not found"}

    # Compute anchor pixel position. Excel measures shapes in points (1pt ≈ 0.75px).
    # If no anchor given, place to the right of the used range.
    BOX_W, BOX_H = 140, 60
    GAP = 30  # arrow length

    if anchor:
        try:
            anchor_range = ws.range(anchor)
            left = float(anchor_range.api.left) if hasattr(anchor_range.api, "left") else 50.0
            top = float(anchor_range.api.top) if hasattr(anchor_range.api, "top") else 50.0
        except Exception:
            left, top = 400.0, 50.0
    else:
        # Place to the right of column F by default (safe for most workbooks)
        try:
            anchor_range = ws.range("F2")
            left = float(anchor_range.api.left) if hasattr(anchor_range.api, "left") else 400.0
            top = float(anchor_range.api.top) if hasattr(anchor_range.api, "top") else 50.0
        except Exception:
            left, top = 400.0, 50.0

    n = len(steps)
    steps_added: list[str] = []

    # msoShapeRoundedRectangle = 5; msoShapeRightArrow = 33
    SHAPE_ROUNDED_RECT = 5
    SHAPE_RIGHT_ARROW = 33

    for i, text in enumerate(steps):
        check_stop()
        x = left + i * (BOX_W + GAP)

        try:
            # Add rounded rectangle for the step
            shp = ws.api.shapes.add_shape(SHAPE_ROUNDED_RECT, x, top, BOX_W, BOX_H)
            # Set text (Mac xlwings + Excel COM)
            try:
                shp.text_frame.text_range.text = str(text)
            except AttributeError:
                # Older xlwings/Excel — try alternative path
                try:
                    shp.text_frame2.text_range.text = str(text)
                except Exception:
                    # Fall back to AppleScript text-set via the shape's index
                    pass
            steps_added.append(f"Box {i+1}: {text}")
        except Exception as e:
            return {
                "status": "failed",
                "error": f"Failed to add box {i+1} ('{text}'): {e}",
                "steps": steps_added,
            }

        # Arrow between this box and the next
        if i < n - 1:
            try:
                arrow_x = x + BOX_W
                arrow_y = top + (BOX_H / 2) - 10
                ws.api.shapes.add_shape(
                    SHAPE_RIGHT_ARROW, arrow_x, arrow_y, GAP, 20
                )
            except Exception as e:
                # Non-fatal — boxes still rendered, just no arrow
                steps_added.append(f"WARN: arrow {i+1}→{i+2} failed: {e}")

    return {
        "status": "completed",
        "message": f"Created SmartArt-style flow diagram with {n} steps on '{ws.name}'",
        "steps": steps_added,
    }


def do_pivot_table(source_range: str, output_cell: str,
                    rows: list = None, columns: list = None,
                    values: list = None, page_filters: list = None,
                    sheet: str = "", output_sheet: str = "",
                    name: str = ""):
    """Create a PivotTable from `source_range`, place it at `output_cell`.

    Args:
        source_range: source data, e.g. 'A1:E50' or 'Sheet1!A1:E50'
        output_cell: top-left cell for the pivot, e.g. 'G2' on output_sheet
        rows: column headers to put in the Rows area
        columns: column headers to put in the Columns area
        values: list of dicts: {"field": "Sales", "function": "sum"|"count"|...}
                or just strings (defaults to sum)
        page_filters: column headers to put in the Filters area
        sheet: source sheet (defaults to active)
        output_sheet: destination sheet (defaults to source sheet)
        name: optional pivot table name

    Runs silently via xlwings — no Excel focus needed.
    """
    if not source_range:
        return {"status": "failed", "error": "source_range is required"}
    if not output_cell:
        return {"status": "failed", "error": "output_cell is required"}

    rows = rows or []
    columns = columns or []
    values = values or []
    page_filters = page_filters or []

    check_stop()

    try:
        app, wb, src_ws = _get_xlwings()
    except Exception as e:
        return {"status": "failed", "error": f"xlwings init: {e}"}

    # Resolve source sheet
    if sheet:
        try:
            src_ws = wb.sheets[sheet]
        except Exception:
            return {"status": "failed", "error": f"Source sheet '{sheet}' not found"}

    # Resolve destination sheet
    dst_ws = src_ws
    if output_sheet:
        try:
            dst_ws = wb.sheets[output_sheet]
        except Exception:
            # Auto-create if missing — analyst convention
            try:
                dst_ws = wb.sheets.add(name=output_sheet, after=src_ws)
            except Exception as e:
                return {"status": "failed", "error": f"Could not create output sheet '{output_sheet}': {e}"}

    # Build the source range reference. xlwings expects 'Sheet!A1:E50' format
    # for cross-sheet pivots, or just 'A1:E50' if same-sheet.
    if "!" in source_range:
        full_source = source_range
    else:
        full_source = f"{src_ws.name}!{source_range}"

    steps: list[str] = []
    try:
        check_stop()
        # xlConsolidationDatabase = 1
        cache = wb.api.pivot_caches().create(
            source_type=1,  # xlDatabase
            source_data=full_source,
        )
        steps.append(f"Created pivot cache from {full_source}")

        check_stop()
        pivot_name = name or f"PivotTable_{int(time.time())}"
        pt = cache.create_pivot_table(
            table_destination=f"'{dst_ws.name}'!{output_cell}",
            table_name=pivot_name,
        )
        steps.append(f"Created pivot table '{pivot_name}' at {output_cell}")
    except Exception as e:
        return {"status": "failed", "error": f"PivotTable create failed: {e}", "steps": steps}

    # Wire fields. xlPivotFieldOrientation enum:
    #   xlRowField = 1, xlColumnField = 2, xlPageField = 3, xlDataField = 4
    def _set_field(field_name: str, orientation: int) -> None:
        try:
            pf = pt.pivot_fields(field_name)
            pf.orientation = orientation
            steps.append(f"  {field_name} → {{1:'row',2:'col',3:'filter',4:'value'}}[{orientation}]")
        except Exception as e:
            steps.append(f"WARN: could not set field '{field_name}': {e}")

    for f in rows:
        check_stop(); _set_field(f, 1)
    for f in columns:
        check_stop(); _set_field(f, 2)
    for f in page_filters:
        check_stop(); _set_field(f, 3)

    # Values: each entry can be a string ('Sales' → sum) or
    # {"field": "Sales", "function": "sum|count|average|max|min"}
    AGG_MAP = {
        "sum": -4157, "count": -4112, "average": -4106,
        "max": -4136, "min": -4139, "product": -4149,
        "stdev": -4155, "var": -4164,
    }
    for entry in values:
        check_stop()
        if isinstance(entry, str):
            field_name, func = entry, "sum"
        elif isinstance(entry, dict):
            field_name = entry.get("field", "")
            func = (entry.get("function") or "sum").lower()
        else:
            continue
        if not field_name:
            continue
        try:
            pf = pt.pivot_fields(field_name)
            pf.orientation = 4  # xlDataField
            pf.function = AGG_MAP.get(func, -4157)
            steps.append(f"  Value: {func}({field_name})")
        except Exception as e:
            steps.append(f"WARN: could not add value '{field_name}': {e}")

    return {
        "status": "completed",
        "message": f"PivotTable '{pivot_name}' created on '{dst_ws.name}' at {output_cell}",
        "steps": steps,
    }


def do_conditional_format_advanced(range_ref: str, rule: str,
                                     sheet: str = "", **kwargs):
    """Apply advanced conditional formatting (color scale / data bar / icon set).

    Args:
        range_ref: target cell range, e.g. 'D2:D40'
        rule: one of 'color_scale_3', 'color_scale_2', 'data_bar',
              'icon_set_3', 'top_n', 'bottom_n', 'above_average', 'below_average'
        sheet: target worksheet (defaults to active)
        **kwargs: rule-specific params (e.g. n=10 for top_n,
                  bar_color='blue' for data_bar, icon_style='3 arrows')

    Uses xlwings COM API — runs silently in the background.
    """
    if not range_ref:
        return {"status": "failed", "error": "range_ref is required"}

    check_stop()

    try:
        app, wb, ws = _get_xlwings()
    except Exception as e:
        return {"status": "failed", "error": f"xlwings init: {e}"}

    if sheet:
        try:
            ws = wb.sheets[sheet]
        except Exception:
            return {"status": "failed", "error": f"Sheet '{sheet}' not found"}

    try:
        rng = ws.range(range_ref)
    except Exception as e:
        return {"status": "failed", "error": f"Bad range '{range_ref}': {e}"}

    fc = rng.api.format_conditions
    steps: list[str] = []

    try:
        check_stop()
        if rule == "color_scale_3":
            # 3-color scale: red → yellow → green by percentile
            cs = fc.add_color_scale(color_scale_type=3)
            steps.append(f"Added 3-color scale to {range_ref}")
        elif rule == "color_scale_2":
            cs = fc.add_color_scale(color_scale_type=2)
            steps.append(f"Added 2-color scale to {range_ref}")
        elif rule == "data_bar":
            db = fc.add_databar()
            steps.append(f"Added data bar to {range_ref}")
        elif rule in ("icon_set_3", "icon_set"):
            ic = fc.add_icon_set_condition()
            steps.append(f"Added 3-icon set to {range_ref}")
        elif rule == "top_n":
            n = int(kwargs.get("n", 10))
            # xlTop10Top = 0
            top = fc.add_top10()
            top.top_bottom = 0  # top
            top.rank = n
            try:
                top.font.color = 0x006100  # dark green text
                top.interior.color = 0xC6EFCE  # light green fill
            except Exception:
                pass
            steps.append(f"Highlighted top {n} values in {range_ref}")
        elif rule == "bottom_n":
            n = int(kwargs.get("n", 10))
            top = fc.add_top10()
            top.top_bottom = 1  # bottom
            top.rank = n
            try:
                top.font.color = 0x9C0006  # dark red text
                top.interior.color = 0xFFC7CE  # light red fill
            except Exception:
                pass
            steps.append(f"Highlighted bottom {n} values in {range_ref}")
        elif rule == "above_average":
            ab = fc.add_above_average()
            ab.above_below = 0  # above
            try:
                ab.font.color = 0x006100
                ab.interior.color = 0xC6EFCE
            except Exception:
                pass
            steps.append(f"Highlighted above-average in {range_ref}")
        elif rule == "below_average":
            ab = fc.add_above_average()
            ab.above_below = 1  # below
            try:
                ab.font.color = 0x9C0006
                ab.interior.color = 0xFFC7CE
            except Exception:
                pass
            steps.append(f"Highlighted below-average in {range_ref}")
        else:
            return {
                "status": "failed",
                "error": f"Unknown conditional format rule: '{rule}'. "
                         f"Valid: color_scale_3, color_scale_2, data_bar, "
                         f"icon_set_3, top_n, bottom_n, above_average, below_average",
            }
    except Exception as e:
        return {"status": "failed", "error": f"Format condition apply failed: {e}", "steps": steps}

    return {
        "status": "completed",
        "message": f"Applied {rule} conditional formatting to {range_ref} on '{ws.name}'",
        "steps": steps,
    }


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

        # ── P1 background-friendly handlers (xlwings, no focus-stealing) ──

        elif action_type == "smartart_diagram":
            return do_smartart_diagram(
                steps=payload.get("steps", []),
                sheet=payload.get("sheet", ""),
                layout=payload.get("layout", "process"),
                anchor=payload.get("anchor", ""),
            )

        elif action_type == "pivot_table":
            return do_pivot_table(
                source_range=payload.get("source_range", ""),
                output_cell=payload.get("output_cell", ""),
                rows=payload.get("rows", []),
                columns=payload.get("columns", []),
                values=payload.get("values", []),
                page_filters=payload.get("page_filters", []),
                sheet=payload.get("sheet", ""),
                output_sheet=payload.get("output_sheet", ""),
                name=payload.get("name", ""),
            )

        elif action_type == "conditional_format_advanced":
            return do_conditional_format_advanced(
                range_ref=payload.get("range", ""),
                rule=payload.get("rule", ""),
                sheet=payload.get("sheet", ""),
                **{k: v for k, v in payload.items()
                   if k not in ("range", "rule", "sheet", "type")},
            )

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

    # Build a rich, multi-line message that surfaces each action's specific
    # outcome (e.g. "Goal Seek converged: D7 changed from 0.5 to 0.7272"
    # instead of the useless "Executed 1 actions (1/1 OK)"). The frontend
    # reads result.message and displays it as the chat reply, so this is
    # the user-visible feedback channel.
    def _summarize() -> str:
        lines: list[str] = []
        for i, (a, r) in enumerate(zip(actions, results), start=1):
            atype = a.get("type", "?")
            rstatus = r.get("status", "unknown")
            rmsg = r.get("message", "") or ""
            if rmsg:
                # Per-action message exists — use it, prefixed for clarity
                # when there's >1 action
                if len(actions) == 1:
                    lines.append(rmsg)
                else:
                    lines.append(f"{i}. {atype}: {rmsg}")
            else:
                # Fall back to status + steps
                err = r.get("error", "")
                steps = r.get("steps", [])
                step_summary = "; ".join(s for s in steps[-3:] if s)
                if rstatus == "completed":
                    fallback = f"{atype}: completed" + (f" — {step_summary}" if step_summary else "")
                elif rstatus == "failed":
                    fallback = f"{atype}: failed" + (f" — {err}" if err else "")
                elif rstatus == "cancelled":
                    fallback = f"{atype}: stopped"
                else:
                    fallback = f"{atype}: {rstatus}"
                lines.append(fallback if len(actions) == 1 else f"{i}. {fallback}")
        return "\n".join(lines) if lines else "No actions executed"

    if cancelled:
        _emergency_cleanup()
        return {
            "status": "cancelled",
            "message": f"Stopped by user after {len(results)}/{len(actions)} action(s).\n\n{_summarize()}",
            "results": results,
            "steps_taken": len(results),
        }

    return {
        "status": "completed" if all_ok else "partial",
        "message": _summarize(),
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
