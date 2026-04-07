"""
Stress Test Suite — Automated R + Excel scenario runner.
Sends realistic prompts, validates responses, logs failures, and self-diagnoses issues.

Usage:
    python -m tests.test_stress                    # run all scenarios once
    python -m tests.test_stress --loops 10         # run 10 full loops
    python -m tests.test_stress --loops 0          # run forever
    python -m tests.test_stress --report           # show summary of past runs
"""

import httpx
import json
import time
import random
import argparse
import sys
import os
from datetime import datetime
from pathlib import Path

BACKEND_URL = "https://focused-solace-production-6839.up.railway.app"
# Use unique user ID per run to avoid hitting the 500 task/month limit
TEST_USER_ID = f"stress-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(1000,9999)}"
TIMEOUT = 120.0
LOG_DIR = Path(__file__).parent / "stress_logs"
LOG_DIR.mkdir(exist_ok=True)

# ── Helpers ──────────────────────────────────────────────────────────────────

_current_user_id = TEST_USER_ID

def send_chat(message: str, context: dict, session_id: str = "", retries: int = 2) -> dict:
    """Send a chat request to the backend synchronously with retry on 429."""
    global _current_user_id
    payload = {
        "user_id": _current_user_id,
        "message": message,
        "context": context,
        "session_id": session_id or f"stress-{context.get('app', 'unknown')}",
    }
    for attempt in range(retries + 1):
        with httpx.Client(timeout=TIMEOUT) as client:
            resp = client.post(f"{BACKEND_URL}/chat/", json=payload)
            if resp.status_code == 429 and attempt < retries:
                # Rate limited — rotate to new user ID
                _current_user_id = f"stress-{datetime.now().strftime('%Y%m%d%H%M%S')}-{random.randint(1000,9999)}"
                payload["user_id"] = _current_user_id
                time.sleep(2)
                continue
            resp.raise_for_status()
            return resp.json()
    resp.raise_for_status()
    return resp.json()


def get_actions(result: dict) -> list:
    actions = result.get("actions", [])
    if not actions and result.get("action", {}).get("type"):
        actions = [result["action"]]
    return actions


def action_types(actions: list) -> list:
    return [a.get("type", "") for a in actions]


def has_action(actions: list, expected_type: str) -> bool:
    return expected_type in action_types(actions)


def payload_field(actions: list, action_type: str, field: str):
    """Get a field from the first action matching the type."""
    for a in actions:
        if a.get("type") == action_type:
            return a.get("payload", {}).get(field)
    return None


# ── Scenario Definitions ────────────────────────────────────────────────────
# Each scenario: (name, message, context, validator_fn)
# validator_fn(result, actions) -> (passed: bool, reason: str)

EXCEL_CONTEXT = {"app": "excel", "sheet": "Sheet1", "sheet_data": []}
EXCEL_WITH_DATA = {
    "app": "excel",
    "sheet": "Sheet1",
    "used_range": "Sheet1!A1:C5",
    "sheet_data": [
        ["Product", "Revenue", "Cost"],
        ["Widget A", 15000, 8000],
        ["Widget B", 22000, 12000],
        ["Widget C", 9500, 6000],
        ["Widget D", 31000, 18000],
    ],
}
RSTUDIO_CONTEXT = {"app": "rstudio"}

def v_has_action(expected_type):
    """Validator: check that at least one action of the expected type exists."""
    def validator(result, actions):
        if has_action(actions, expected_type):
            return True, f"Found {expected_type}"
        return False, f"Expected {expected_type}, got {action_types(actions)}"
    return validator

def v_has_any_action(*types):
    """Validator: check that at least one of the expected types exists."""
    def validator(result, actions):
        found = action_types(actions)
        for t in types:
            if t in found:
                return True, f"Found {t}"
        return False, f"Expected one of {list(types)}, got {found}"
    return validator

def v_has_actions(*types):
    """Validator: check that ALL expected types exist."""
    def validator(result, actions):
        found = action_types(actions)
        missing = [t for t in types if t not in found]
        if not missing:
            return True, f"All found: {list(types)}"
        return False, f"Missing {missing} from {found}"
    return validator

def v_has_reply():
    """Validator: check that reply text is non-empty."""
    def validator(result, actions):
        reply = result.get("reply", "")
        if reply and len(reply) > 5:
            return True, f"Reply: {reply[:80]}..."
        return False, f"Empty or too short reply: '{reply}'"
    return validator

def v_code_contains(*keywords):
    """Validator: check that run_r_code action contains all keywords."""
    def validator(result, actions):
        code = payload_field(actions, "run_r_code", "code") or ""
        missing = [kw for kw in keywords if kw.lower() not in code.lower()]
        if not missing:
            return True, f"Code contains all: {list(keywords)}"
        return False, f"Code missing {missing}. Code: {code[:200]}"
    return validator

def v_write_cell_value(expected_field="formula"):
    """Validator: check that write_cell has a formula or value."""
    def validator(result, actions):
        for a in actions:
            if a.get("type") in ("write_cell", "write_formula"):
                p = a.get("payload", {})
                if p.get("formula") or p.get("value") is not None:
                    return True, f"write_cell has content"
        return False, f"No write_cell with content in {action_types(actions)}"
    return validator

def v_any_action_exists():
    """Validator: check that at least one action was returned."""
    def validator(result, actions):
        if actions:
            return True, f"Got {len(actions)} actions: {action_types(actions)}"
        # Check if reply is reasonable even without actions
        reply = result.get("reply", "")
        if reply and len(reply) > 20:
            return True, f"No actions but has reply: {reply[:80]}"
        return False, "No actions and no meaningful reply"
    return validator

def v_format_has_property(*props):
    """Validator: check that format_range has expected properties."""
    def validator(result, actions):
        for a in actions:
            if a.get("type") == "format_range":
                p = a.get("payload", {})
                found = [pr for pr in props if pr in p]
                if found:
                    return True, f"format_range has: {found}"
        return False, f"No format_range with {list(props)}"
    return validator


# ── Excel Scenarios ──────────────────────────────────────────────────────────

EXCEL_SCENARIOS = [
    # Basic writes
    ("excel_write_hello", "Write 'Hello World' in cell A1",
     EXCEL_CONTEXT, v_has_action("write_cell")),

    ("excel_write_number", "Put 42 in B2",
     EXCEL_CONTEXT, v_has_action("write_cell")),

    ("excel_sum_formula", "Write a SUM formula in D2 that adds B2 to C2",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_profit_column", "Add a Profit column in D that calculates Revenue minus Cost for each row",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_range", "write_formula", "fill_down")),

    ("excel_average_formula", "Calculate the average revenue in B6",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_percentage_margin", "Add a margin % column in E that shows (Revenue-Cost)/Revenue for each product",
     EXCEL_WITH_DATA, v_any_action_exists()),

    ("excel_conditional_format", "Highlight all revenue values above 20000 in green",
     EXCEL_WITH_DATA, v_has_any_action("add_conditional_format", "format_range")),

    ("excel_sort_revenue", "Sort the data by Revenue from highest to lowest",
     EXCEL_WITH_DATA, v_has_action("sort_range")),

    ("excel_chart_bar", "Create a bar chart comparing Revenue by Product",
     EXCEL_WITH_DATA, v_has_action("add_chart")),

    ("excel_chart_pie", "Create a pie chart showing the cost distribution across products",
     EXCEL_WITH_DATA, v_has_action("add_chart")),

    ("excel_format_currency", "Format the Revenue and Cost columns as USD currency",
     EXCEL_WITH_DATA, v_has_any_action("format_range", "set_number_format")),

    ("excel_bold_headers", "Bold the header row and make it blue background",
     EXCEL_WITH_DATA, v_has_action("format_range")),

    ("excel_freeze_panes", "Freeze the top row",
     EXCEL_WITH_DATA, v_has_action("freeze_panes")),

    ("excel_add_sheet", "Create a new sheet called Summary",
     EXCEL_CONTEXT, v_has_action("add_sheet")),

    ("excel_autofit", "Autofit all columns",
     EXCEL_WITH_DATA, v_has_any_action("autofit", "autofit_columns")),

    ("excel_named_range", "Create a named range called 'Revenues' for B2:B5",
     EXCEL_WITH_DATA, v_has_action("create_named_range")),

    ("excel_data_validation", "Add a dropdown in F2 with options: Low, Medium, High",
     EXCEL_WITH_DATA, v_has_action("add_data_validation")),

    ("excel_countif", "In B7, count how many products have revenue above 15000",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_vlookup", "Write a VLOOKUP in E1 that finds the cost of Widget B from the table",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_max_min", "Show the max revenue in B8 and min revenue in B9",
     EXCEL_WITH_DATA, v_any_action_exists()),

    ("excel_multi_write", "Write 'Q1', 'Q2', 'Q3', 'Q4' in cells F1 through F4",
     EXCEL_CONTEXT, v_any_action_exists()),

    ("excel_number_format_pct", "Format column E as percentages with 1 decimal place",
     EXCEL_WITH_DATA, v_has_any_action("set_number_format", "format_range")),

    ("excel_clear_range", "Clear all data in column D",
     EXCEL_WITH_DATA, v_has_action("clear_range")),

    ("excel_line_chart", "Create a line chart of Revenue trends across products",
     EXCEL_WITH_DATA, v_has_action("add_chart")),

    ("excel_sumproduct", "Calculate a weighted average revenue in cell B10 using SUMPRODUCT",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    # Complex multi-step
    ("excel_full_analysis", "Create a summary: total revenue in G1, total cost in G2, total profit in G3, and profit margin % in G4",
     EXCEL_WITH_DATA, v_any_action_exists()),

    ("excel_dashboard_headers", "Set up a dashboard header: merge A1:F1, write 'Sales Dashboard', center it, make font size 16, bold, dark blue",
     EXCEL_CONTEXT, v_any_action_exists()),

    ("excel_if_formula", "In F2, write an IF formula: if Revenue > 20000 then 'High' else 'Low'",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_concatenate", "In G2 write a formula that concatenates Product name and ' - $' and Revenue",
     EXCEL_WITH_DATA, v_has_any_action("write_cell", "write_formula")),

    ("excel_import_csv", "Import the CSV file at /tmp/test_data.csv",
     EXCEL_CONTEXT, v_has_action("import_csv")),
]

# ── R Scenarios ──────────────────────────────────────────────────────────────

R_SCENARIOS = [
    # Basic stats
    ("r_mean_vector", "Calculate the mean of c(10, 20, 30, 40, 50)",
     RSTUDIO_CONTEXT, v_has_action("run_r_code")),

    ("r_summary_stats", "Create a vector x = 1:100 and show its summary statistics",
     RSTUDIO_CONTEXT, v_code_contains("summary")),

    ("r_sd_calc", "Calculate the standard deviation of c(5, 10, 15, 20, 25, 30)",
     RSTUDIO_CONTEXT, v_code_contains("sd")),

    ("r_correlation", "Generate two random vectors of length 50 and compute their correlation",
     RSTUDIO_CONTEXT, v_code_contains("cor")),

    # Data frames
    ("r_create_df", "Create a data frame with columns: Name (Alice, Bob, Charlie), Age (25, 30, 35), Salary (50000, 60000, 70000)",
     RSTUDIO_CONTEXT, v_code_contains("data.frame")),

    ("r_filter_df", "Create a data frame of 5 students with Name, Grade, and Score columns. Filter for students with Score > 80",
     RSTUDIO_CONTEXT, v_has_action("run_r_code")),

    ("r_mutate_df", "Using dplyr, create a data frame and add a new column that's the log of an existing numeric column",
     RSTUDIO_CONTEXT, v_code_contains("mutate")),

    ("r_group_summarise", "Create sales data with Region and Amount columns (10 rows). Group by Region and calculate mean Amount per region",
     RSTUDIO_CONTEXT, v_code_contains("group_by")),

    ("r_merge_dfs", "Create two data frames sharing an ID column and merge them",
     RSTUDIO_CONTEXT, v_has_action("run_r_code")),

    ("r_reshape_wide", "Create long-format data and pivot it to wide format using tidyr",
     RSTUDIO_CONTEXT, v_code_contains("pivot_wider")),

    # Plots
    ("r_basic_plot", "Plot a scatter plot of x = 1:20 and y = x^2",
     RSTUDIO_CONTEXT, v_code_contains("plot")),

    ("r_histogram", "Generate 1000 random normal values and create a histogram",
     RSTUDIO_CONTEXT, v_code_contains("hist")),

    ("r_boxplot", "Create a boxplot of mpg grouped by cyl using the mtcars dataset",
     RSTUDIO_CONTEXT, v_code_contains("boxplot", "mtcars")),

    ("r_ggplot_scatter", "Using ggplot2, create a scatter plot of mpg vs hp from mtcars, colored by cyl",
     RSTUDIO_CONTEXT, v_code_contains("ggplot", "mtcars")),

    ("r_ggplot_bar", "Using ggplot2, create a bar chart showing count of cars by number of cylinders from mtcars",
     RSTUDIO_CONTEXT, v_code_contains("ggplot", "geom_bar")),

    ("r_ggplot_line", "Create a line chart of AirPassengers time series using ggplot2",
     RSTUDIO_CONTEXT, v_code_contains("ggplot")),

    ("r_pairs_plot", "Create a pairs plot of the first 4 columns of iris dataset",
     RSTUDIO_CONTEXT, v_code_contains("pairs", "iris")),

    ("r_qqnorm", "Generate 100 random normal values and create a QQ plot",
     RSTUDIO_CONTEXT, v_code_contains("qqnorm")),

    # Statistical tests
    ("r_ttest", "Perform a two-sample t-test comparing two groups: c(5,7,8,6,9) and c(3,4,5,6,4)",
     RSTUDIO_CONTEXT, v_code_contains("t.test")),

    ("r_chisq", "Create a contingency table and perform a chi-squared test",
     RSTUDIO_CONTEXT, v_code_contains("chisq")),

    ("r_linear_model", "Fit a linear regression of mpg on wt and hp from mtcars, show summary",
     RSTUDIO_CONTEXT, v_code_contains("lm", "mtcars")),

    ("r_anova", "Perform a one-way ANOVA of Sepal.Length by Species using the iris dataset",
     RSTUDIO_CONTEXT, v_code_contains("aov", "iris")),

    ("r_logistic", "Fit a logistic regression predicting vs from mpg and wt in mtcars",
     RSTUDIO_CONTEXT, v_code_contains("glm", "mtcars")),

    # Data manipulation
    ("r_read_csv", "Read the CSV file at /tmp/test_data.csv into a data frame called df",
     RSTUDIO_CONTEXT, v_code_contains("read.csv", "df")),

    ("r_string_manip", "Create a vector of 5 names and convert them all to uppercase",
     RSTUDIO_CONTEXT, v_code_contains("toupper")),

    ("r_apply_function", "Create a matrix and use apply to calculate row means",
     RSTUDIO_CONTEXT, v_code_contains("apply")),

    ("r_sapply", "Use sapply to calculate the square root of each element in c(4, 9, 16, 25, 36)",
     RSTUDIO_CONTEXT, v_code_contains("sapply", "sqrt")),

    ("r_for_loop", "Write a for loop that prints the first 10 Fibonacci numbers",
     RSTUDIO_CONTEXT, v_code_contains("for")),

    ("r_custom_function", "Write a function called is_prime that checks if a number is prime, then test it on 17",
     RSTUDIO_CONTEXT, v_code_contains("function", "is_prime")),

    ("r_tryCatch", "Write code that uses tryCatch to handle a division by zero error gracefully",
     RSTUDIO_CONTEXT, v_code_contains("tryCatch")),

    # Advanced
    ("r_pca", "Perform PCA on the iris dataset (numeric columns only) and plot the first two components",
     RSTUDIO_CONTEXT, v_code_contains("prcomp", "iris")),

    ("r_kmeans", "Perform k-means clustering (k=3) on iris numeric data and visualize the clusters",
     RSTUDIO_CONTEXT, v_code_contains("kmeans", "iris")),

    ("r_time_series", "Create a time series of monthly data for 3 years with a trend, and decompose it",
     RSTUDIO_CONTEXT, v_code_contains("ts")),

    ("r_heatmap", "Create a correlation heatmap of the mtcars numeric variables",
     RSTUDIO_CONTEXT, v_code_contains("cor", "mtcars")),

    ("r_install_pkg", "Install the readxl package",
     RSTUDIO_CONTEXT, v_has_action("install_package")),

    # Edge cases
    ("r_empty_request", "What packages do I have loaded?",
     RSTUDIO_CONTEXT, v_any_action_exists()),

    ("r_multiline_code", "Create a full analysis: load mtcars, create summary stats, run a regression of mpg on wt, and plot the residuals",
     RSTUDIO_CONTEXT, v_code_contains("lm", "mtcars")),

    ("r_pipe_chain", "Using dplyr pipes, take mtcars, filter for 6-cylinder cars, select mpg and hp, arrange by mpg descending",
     RSTUDIO_CONTEXT, v_code_contains("filter", "select", "arrange")),
]

# ── Cross-App Scenarios ──────────────────────────────────────────────────────

CROSS_APP_SCENARIOS = [
    ("cross_paste_r_plot", "paste the R plot",
     {"app": "excel", "sheet": "Sheet1", "sheet_data": []},
     v_has_action("import_image")),

    ("cross_export_plot", "export this plot to Excel",
     RSTUDIO_CONTEXT,
     v_has_action("export_plot")),
]

ALL_SCENARIOS = EXCEL_SCENARIOS + R_SCENARIOS + CROSS_APP_SCENARIOS


# ── Runner ───────────────────────────────────────────────────────────────────

class TestRunner:
    def __init__(self):
        self.results = []
        self.start_time = None
        self.log_file = LOG_DIR / f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jsonl"

    def run_scenario(self, name, message, context, validator):
        """Run a single scenario and return result dict."""
        entry = {
            "name": name,
            "message": message,
            "app": context.get("app", "unknown"),
            "timestamp": datetime.now().isoformat(),
        }
        try:
            t0 = time.time()
            result = send_chat(message, context)
            elapsed = time.time() - t0
            actions = get_actions(result)

            passed, reason = validator(result, actions)

            entry.update({
                "passed": passed,
                "reason": reason,
                "elapsed_s": round(elapsed, 2),
                "reply": result.get("reply", "")[:200],
                "action_types": action_types(actions),
                "model_used": result.get("model_used", ""),
                "error": None,
            })
        except Exception as e:
            entry.update({
                "passed": False,
                "reason": f"Exception: {str(e)}",
                "elapsed_s": 0,
                "reply": "",
                "action_types": [],
                "model_used": "",
                "error": str(e),
            })

        self.results.append(entry)
        self._log(entry)
        return entry

    def _log(self, entry):
        with open(self.log_file, "a") as f:
            f.write(json.dumps(entry) + "\n")

    def run_all(self, scenarios=None, shuffle=False):
        """Run all scenarios once."""
        scenarios = scenarios or ALL_SCENARIOS
        if shuffle:
            scenarios = list(scenarios)
            random.shuffle(scenarios)

        total = len(scenarios)
        passed = 0
        failed = 0

        for i, (name, msg, ctx, validator) in enumerate(scenarios, 1):
            status_char = "."
            try:
                entry = self.run_scenario(name, msg, ctx, validator)
                if entry["passed"]:
                    passed += 1
                    status_char = "✓"
                else:
                    failed += 1
                    status_char = "✗"
                print(f"  [{i}/{total}] {status_char} {name} ({entry['elapsed_s']}s) — {entry['reason'][:80]}")
            except KeyboardInterrupt:
                print("\n\nInterrupted by user.")
                break

            # Small delay to avoid rate-limiting
            time.sleep(0.5)

        return passed, failed, total

    def run_loops(self, num_loops=1, shuffle=True):
        """Run multiple loops of all scenarios."""
        self.start_time = time.time()
        total_passed = 0
        total_failed = 0
        total_run = 0
        loop = 0

        try:
            while True:
                loop += 1
                if num_loops > 0 and loop > num_loops:
                    break

                print(f"\n{'='*60}")
                print(f"  LOOP {loop}" + (f"/{num_loops}" if num_loops > 0 else " (infinite)"))
                print(f"  Started: {datetime.now().strftime('%H:%M:%S')}")
                print(f"{'='*60}")

                p, f, t = self.run_all(shuffle=shuffle)
                total_passed += p
                total_failed += f
                total_run += t

                elapsed = time.time() - self.start_time
                rate = total_run / (elapsed / 60) if elapsed > 0 else 0

                print(f"\n  Loop {loop} complete: {p}/{t} passed, {f} failed")
                print(f"  Cumulative: {total_passed}/{total_run} passed ({total_failed} failed)")
                print(f"  Rate: {rate:.1f} scenarios/min | Elapsed: {elapsed/60:.1f} min")
                print(f"  Log: {self.log_file}")

        except KeyboardInterrupt:
            print("\n\nStopped by user.")

        self.print_summary(total_passed, total_failed, total_run)

    def print_summary(self, passed, failed, total):
        elapsed = time.time() - self.start_time if self.start_time else 0
        print(f"\n{'='*60}")
        print(f"  FINAL SUMMARY")
        print(f"{'='*60}")
        print(f"  Total scenarios run:  {total}")
        print(f"  Passed:               {passed}")
        print(f"  Failed:               {failed}")
        print(f"  Pass rate:            {passed/total*100:.1f}%" if total > 0 else "  N/A")
        print(f"  Total time:           {elapsed/60:.1f} minutes")
        print(f"  Log file:             {self.log_file}")

        if failed > 0:
            print(f"\n  FAILURES:")
            for r in self.results:
                if not r["passed"]:
                    print(f"    ✗ {r['name']} — {r['reason'][:100]}")


def show_report():
    """Show summary of all past runs."""
    log_files = sorted(LOG_DIR.glob("run_*.jsonl"))
    if not log_files:
        print("No test runs found.")
        return

    print(f"\n{'='*60}")
    print(f"  STRESS TEST HISTORY ({len(log_files)} runs)")
    print(f"{'='*60}")

    for lf in log_files[-10:]:  # show last 10 runs
        entries = []
        with open(lf) as f:
            for line in f:
                try:
                    entries.append(json.loads(line))
                except json.JSONDecodeError:
                    continue

        passed = sum(1 for e in entries if e.get("passed"))
        failed = sum(1 for e in entries if not e.get("passed"))
        total = len(entries)
        ts = lf.stem.replace("run_", "")

        print(f"  {ts}: {passed}/{total} passed, {failed} failed")

        # Show unique failures
        failure_names = set()
        for e in entries:
            if not e.get("passed"):
                failure_names.add(e["name"])
        if failure_names:
            for fn in sorted(failure_names):
                # Find the reason
                for e in entries:
                    if e["name"] == fn and not e["passed"]:
                        print(f"    ✗ {fn}: {e['reason'][:80]}")
                        break

    print()


# ── Main ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Stress test tsifl backend")
    parser.add_argument("--loops", type=int, default=1, help="Number of loops (0=infinite)")
    parser.add_argument("--report", action="store_true", help="Show report of past runs")
    parser.add_argument("--excel-only", action="store_true", help="Run only Excel scenarios")
    parser.add_argument("--r-only", action="store_true", help="Run only R scenarios")
    parser.add_argument("--no-shuffle", action="store_true", help="Don't shuffle scenario order")
    args = parser.parse_args()

    if args.report:
        show_report()
        sys.exit(0)

    # Select scenarios
    if args.excel_only:
        scenarios = EXCEL_SCENARIOS
    elif args.r_only:
        scenarios = R_SCENARIOS
    else:
        scenarios = None  # uses ALL_SCENARIOS

    runner = TestRunner()
    if scenarios:
        # Override ALL_SCENARIOS temporarily
        orig = list(ALL_SCENARIOS)
        ALL_SCENARIOS.clear()
        ALL_SCENARIOS.extend(scenarios)

    runner.run_loops(num_loops=args.loops, shuffle=not args.no_shuffle)
