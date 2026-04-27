"""
Local services for the desktop agent.
Builds natural language instructions from structured action payloads.
"""

import json


def build_instructions(actions: list, context: dict) -> str:
    """Convert structured actions into natural language instructions
    that Claude can follow while controlling the screen."""

    parts = [
        "You are controlling a Mac with Microsoft Excel open.",
        "The user's Excel workbook is already open and visible on screen.",
        "",
        "IMPORTANT RULES:",
        "- Take a screenshot FIRST to see the current state",
        "- After each action, take a screenshot to verify it worked",
        "- Use the Excel ribbon menus (Data tab) for What-If Analysis features",
        "- Click precisely on menu items, buttons, and dialog fields",
        "- If a dialog appears, fill in all fields before clicking OK",
        "- On macOS, use Cmd instead of Ctrl for keyboard shortcuts",
        "",
        "Perform these tasks IN ORDER:",
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

            parts.append(f"TASK {i}: Create a Data Table")
            if sheet:
                parts.append(f"  1. Click on the '{sheet}' sheet tab at the bottom")
            parts.append(f"  2. Select the range {table_range}")
            parts.append(f"     (Click the first cell, hold Shift, click the last cell)")
            parts.append(f"  3. Click the 'Data' tab in the ribbon")
            parts.append(f"  4. Click 'What-If Analysis' dropdown")
            parts.append(f"  5. Click 'Data Table...'")
            if row_input and col_input:
                parts.append(f"  6. In the dialog:")
                parts.append(f"     - Row input cell: type {row_input}")
                parts.append(f"     - Column input cell: type {col_input}")
            elif col_input:
                parts.append(f"  6. In the dialog:")
                parts.append(f"     - Leave 'Row input cell' empty")
                parts.append(f"     - Column input cell: type {col_input}")
            elif row_input:
                parts.append(f"  6. In the dialog:")
                parts.append(f"     - Row input cell: type {row_input}")
                parts.append(f"     - Leave 'Column input cell' empty")
            parts.append(f"  7. Click OK")

        elif action_type == "scenario_manager":
            name = payload.get("name", "")
            cells = payload.get("changing_cells", "")
            values = payload.get("values", [])

            parts.append(f"TASK {i}: Create Scenario '{name}'")
            parts.append(f"  1. Click the 'Data' tab in the ribbon")
            parts.append(f"  2. Click 'What-If Analysis' dropdown")
            parts.append(f"  3. Click 'Scenario Manager...'")
            parts.append(f"  4. Click 'Add...'")
            parts.append(f"  5. Scenario name: type '{name}'")
            parts.append(f"  6. Changing cells: type {cells}")
            parts.append(f"  7. Click OK")
            if values:
                parts.append(f"  8. Enter values: {', '.join(str(v) for v in values)}")
            parts.append(f"  9. Click OK")
            parts.append(f"  10. Click 'Close' in Scenario Manager")

        elif action_type == "save_solver_scenario":
            name = payload.get("name", "Solver Solution")
            objective = payload.get("objective_cell", "")
            goal = payload.get("goal", "max")
            changing = payload.get("changing_cells", "")
            constraints = payload.get("constraints", [])

            parts.append(f"TASK {i}: Run Solver and Save as Scenario '{name}'")
            parts.append(f"  1. Click the 'Data' tab in the ribbon")
            parts.append(f"  2. Click 'Solver' (far right of Data tab)")
            parts.append(f"  3. Set Objective: {objective}")
            if goal == "max":
                parts.append(f"  4. Select 'Max'")
            elif goal == "min":
                parts.append(f"  4. Select 'Min'")
            else:
                parts.append(f"  4. Select 'Value Of:' and type {goal}")
            parts.append(f"  5. By Changing Variable Cells: {changing}")
            for j, c in enumerate(constraints):
                parts.append(f"  6.{j+1}. Add constraint: {c}")
            parts.append(f"  7. Click 'Solve'")
            parts.append(f"  8. In results dialog, click 'Save Scenario...'")
            parts.append(f"  9. Name: '{name}'")
            parts.append(f"  10. Click OK, then OK again")

        elif action_type == "run_solver":
            objective = payload.get("objective_cell", "")
            goal = payload.get("goal", "max")
            changing = payload.get("changing_cells", "")

            parts.append(f"TASK {i}: Run Solver")
            parts.append(f"  1. Click the 'Data' tab")
            parts.append(f"  2. Click 'Solver'")
            parts.append(f"  3. Set Objective: {objective}, Goal: {goal}")
            parts.append(f"  4. Changing cells: {changing}")
            parts.append(f"  5. Click 'Solve'")
            parts.append(f"  6. Click OK to keep Solver solution")

        elif action_type == "goal_seek":
            parts.append(f"TASK {i}: Goal Seek")
            parts.append(f"  1. Click the 'Data' tab")
            parts.append(f"  2. Click 'What-If Analysis' > 'Goal Seek...'")
            parts.append(f"  3. Set cell: {payload.get('set_cell', '')}")
            parts.append(f"  4. To value: {payload.get('to_value', '')}")
            parts.append(f"  5. By changing cell: {payload.get('changing_cell', '')}")
            parts.append(f"  6. Click OK, then OK again")

        elif action_type == "scenario_summary":
            parts.append(f"TASK {i}: Create Scenario Summary Report")
            parts.append(f"  1. Click the 'Data' tab")
            parts.append(f"  2. Click 'What-If Analysis' > 'Scenario Manager...'")
            parts.append(f"  3. Click 'Summary...'")
            parts.append(f"  4. Result cells: {payload.get('result_cells', '')}")
            parts.append(f"  5. Report type: Scenario summary (should be default)")
            parts.append(f"  6. Click OK")

        elif action_type == "run_toolpak":
            tool = payload.get("tool", "Descriptive Statistics")
            input_range = payload.get("input_range", "")
            output_range = payload.get("output_range", "")
            options = payload.get("options", {})

            parts.append(f"TASK {i}: Run Analysis ToolPak — {tool}")
            parts.append(f"  1. Click the 'Data' tab")
            parts.append(f"  2. Click 'Data Analysis' (far right)")
            parts.append(f"  3. Select '{tool}' from the list")
            parts.append(f"  4. Click OK")
            parts.append(f"  5. Input Range: {input_range}")
            if output_range:
                parts.append(f"  6. Output Range: select 'Output Range' and type {output_range}")
            if options.get("summary_statistics"):
                parts.append(f"  7. Check 'Summary statistics'")
            if options.get("labels_in_first_row"):
                parts.append(f"  8. Check 'Labels in first row'")
            if options.get("grouped_by") == "columns":
                parts.append(f"  9. Grouped By: Columns")
            parts.append(f"  10. Click OK")

        elif action_type == "computer_use":
            # Generic CU fallback — backend Claude emits this for ribbon-only
            # features that don't have a dedicated action type (SmartArt,
            # page setup, headers/footers, etc.). The `task` field carries
            # a step-by-step description that Claude CU can follow directly.
            task = payload.get("task", "").strip()
            if task:
                parts.append(f"TASK {i}: {task}")
            else:
                # Fallback: backend forgot the task field. Dump payload so CU
                # has *something* to work with rather than failing silently.
                parts.append(f"TASK {i}: {json.dumps(payload)}")

        else:
            parts.append(f"TASK {i}: {action_type}")
            parts.append(f"  Details: {json.dumps(payload)}")

        parts.append("")

    parts.append("After completing ALL tasks, confirm what was done.")
    parts.append("IMPORTANT: Take a screenshot first to see the current state of Excel.")

    return "\n".join(parts)
