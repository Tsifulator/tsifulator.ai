#!/usr/bin/env python3
"""
Generate the 20-case hardening sprint regression suite.

Each spec below produces a `cases/<id>/request.json` + `cases/<id>/rubric.yaml`
pair. Cases cover:

  005–007  Goal Seek (the loop we just closed)
  008–012  Core Excel analyst basics (formulas, fill_down, format, charts)
  013–015  Path B P1 handlers (smartart, pivot, conditional formatting)
  016–017  Multi-sheet / cross-sheet patterns
  018–020  Hallucination guards (polish-inject, ## errors, discuss-mode)
  021–023  Edge cases (phantom sheet, locked cell, impossible feature)
  024      RStudio (RMarkdown file generation)

Run with:  python3 _build_hardening_cases.py
"""

import json
import os
from pathlib import Path

import yaml  # if missing: pip3 install pyyaml

CASES_DIR = Path(__file__).parent / "cases"


def _make_excel_context(
    sheet="Sheet1",
    all_sheets=None,
    used_range="A1:E10",
    sheet_data=None,
    sheet_summaries=None,
    workbook_name="Test.xlsx",
):
    """Build a typical Excel context block for chat requests."""
    if all_sheets is None:
        all_sheets = [sheet]
    if sheet_data is None:
        sheet_data = []
    if sheet_summaries is None:
        sheet_summaries = [
            {
                "name": s,
                "used_range": used_range if s == sheet else "empty",
                "rows": len(sheet_data) if s == sheet else 0,
                "cols": (max((len(r) for r in sheet_data), default=0)
                         if s == sheet else 0),
                "preview": sheet_data if s == sheet else [],
                "preview_formulas": sheet_data if s == sheet else [],
            }
            for s in all_sheets
        ]
    return {
        "app": "excel",
        "sheet": sheet,
        "workbook_name": workbook_name,
        "all_sheets": all_sheets,
        "selected_cell": "A1",
        "selected_value": None,
        "used_range": used_range,
        "sheet_data": sheet_data,
        "sheet_formulas": sheet_data,
        "sheet_summaries": sheet_summaries,
        "named_ranges": [],
    }


# ── Case specs ──────────────────────────────────────────────────────────────


CASES = [

    # ── 005 — Goal Seek with explicit cells (the test we just confirmed works) ─
    {
        "id": "005-goal-seek-explicit-cells",
        "request": {
            "user_id": "test-005",
            "message": (
                "On the Calculator sheet, run goal seek to make C18 equal to "
                "100000 by changing D7"
            ),
            "context": _make_excel_context(
                sheet="Calculator",
                all_sheets=["Calculator"],
                used_range="A1:E26",
                sheet_data=[
                    [None, None, None, None, None],   # row 1
                ] + [[None] * 5] * 16 + [             # rows 2-17 padding
                    [None, None, "=C14*D7", None, None],  # row 18 — C18 has formula
                ] + [[None] * 5] * 8,                 # rows 19-26 padding
            ),
        },
        "rubric": {
            "_project": "Goal Seek with explicit set/changing cells",
            "min_action_count": 1,
            "must_include_action_types": ["goal_seek"],
            "must_reference_cells": ["Calculator!D7"],  # changing_cell must be D7
            "reply_not_empty": False,  # CU dispatch blanks reply
        },
    },

    # ── 006 — Polish injector fires on vague "fix" prompt with #### errors ────
    {
        "id": "006-polish-inject-on-pound-errors",
        "request": {
            "user_id": "test-006",
            "message": "fix the #### errors on this sheet",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:D10",
                sheet_data=[
                    ["Item", "Price", "Qty", "Total"],
                    ["Widget A", 1234567.89, 100, 123456789],
                    ["Widget B", 9876543.21, 50, 493827160.5],
                ],
            ),
        },
        "rubric": {
            "_project": "Polish auto-injector on '#### errors' prompt",
            # Polish injector should add autofit even if model stalls
            "min_action_count": 1,
            "must_include_action_types": ["autofit"],
        },
    },

    # ── 007 — Polish injector on "make it look better" ────────────────────────
    {
        "id": "007-polish-inject-make-it-look-better",
        "request": {
            "user_id": "test-007",
            "message": "make this workbook look better",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:C5",
                sheet_data=[
                    ["Name", "Revenue", "Profit"],
                    ["Q1", 100000, 25000],
                    ["Q2", 120000, 30000],
                ],
            ),
        },
        "rubric": {
            "_project": "Polish injector on 'make it look better'",
            "min_action_count": 1,
            # Should emit at least autofit; ideally also format_range / freeze
        },
    },

    # ── 008 — Write formula + fill_down (NEVER WRITE ROW-BY-ROW rule) ─────────
    {
        "id": "008-write-formula-fill-down",
        "request": {
            "user_id": "test-008",
            "message": (
                "On Sheet1, calculate Total = Price * Quantity in column D for "
                "all rows from D2 to D40"
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:D40",
                sheet_data=[
                    ["Item", "Price", "Qty", "Total"],
                    ["Item 1", 10, 5, None],
                    ["Item 2", 20, 3, None],
                ],
            ),
        },
        "rubric": {
            "_project": "Write formula once + fill_down (no row-by-row writes)",
            "min_action_count": 1,
            "must_include_action_types": ["write_formula", "fill_down"],
            # Forbidden: writing the same formula 38 times instead of 1+fill_down
            "max_action_count": 5,
            "formula_at_cell_contains": {
                "Sheet1!D2": ["B2", "C2"],
            },
        },
    },

    # ── 009 — Named range create ──────────────────────────────────────────────
    {
        "id": "009-named-range-create",
        "request": {
            "user_id": "test-009",
            "message": "Create a named range called 'Revenue' for cells B2:B13 on Sheet1",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:B13",
                sheet_data=[
                    ["Month", "Revenue"],
                    ["Jan", 100],
                    ["Feb", 200],
                ],
            ),
        },
        "rubric": {
            "_project": "Named range creation",
            "min_action_count": 1,
            "must_include_action_types": ["create_named_range"],
        },
    },

    # ── 010 — Format currency on a column ─────────────────────────────────────
    {
        "id": "010-format-currency-column",
        "request": {
            "user_id": "test-010",
            "message": "Format column C on Sheet1 as currency with 2 decimals",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:C5",
                sheet_data=[
                    ["Item", "Qty", "Price"],
                    ["Apple", 10, 1.5],
                    ["Banana", 20, 0.75],
                ],
            ),
        },
        "rubric": {
            "_project": "Currency formatting",
            "min_action_count": 1,
            "must_include_action_types": ["format_range"],
        },
    },

    # ── 011 — Add bar chart from a range ──────────────────────────────────────
    {
        "id": "011-add-bar-chart",
        "request": {
            "user_id": "test-011",
            "message": "Add a bar chart on Sheet1 showing revenue by quarter from A1:B5",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:B5",
                sheet_data=[
                    ["Quarter", "Revenue"],
                    ["Q1", 100000],
                    ["Q2", 120000],
                    ["Q3", 95000],
                    ["Q4", 140000],
                ],
            ),
        },
        "rubric": {
            "_project": "Add bar chart",
            "min_action_count": 1,
            "must_include_action_types": ["add_chart"],
        },
    },

    # ── 012 — INDEX/MATCH lookup ──────────────────────────────────────────────
    {
        "id": "012-index-match-lookup",
        "request": {
            "user_id": "test-012",
            "message": (
                "On Sheet1, in cell E2, write a formula that uses INDEX/MATCH "
                "to look up the price in column B based on the item name in D2"
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:E5",
                sheet_data=[
                    ["Item", "Price", None, "Lookup", "Result"],
                    ["Apple", 1.5, None, "Apple", None],
                    ["Banana", 0.75, None, None, None],
                ],
            ),
        },
        "rubric": {
            "_project": "INDEX/MATCH formula",
            "min_action_count": 1,
            "must_include_action_types": ["write_formula"],
            "formula_at_cell_contains": {
                "Sheet1!E2": ["INDEX", "MATCH"],
            },
        },
    },

    # ── 013 — SmartArt P1 handler ─────────────────────────────────────────────
    {
        "id": "013-smartart-diagram-p1",
        "request": {
            "user_id": "test-013",
            "message": (
                "On Sheet1, insert a SmartArt diagram showing the workflow "
                "with 4 steps: Plan, Build, Test, Ship. Use a horizontal "
                "arrow layout."
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:A1",
                sheet_data=[],
            ),
        },
        "rubric": {
            "_project": "SmartArt P1 handler emission",
            "min_action_count": 1,
            "must_include_action_types": ["smartart_diagram"],
        },
    },

    # ── 014 — PivotTable P1 handler ───────────────────────────────────────────
    {
        "id": "014-pivot-table-p1",
        "request": {
            "user_id": "test-014",
            "message": (
                "Create a PivotTable from the data on Sheet1 (A1:D20) with "
                "Region as rows, Product as columns, and sum of Sales as the "
                "value. Place it on a new tab called Pivot."
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:D20",
                sheet_data=[
                    ["Date", "Region", "Product", "Sales"],
                    ["2026-01-01", "North", "A", 100],
                    ["2026-01-02", "South", "B", 200],
                ],
            ),
        },
        "rubric": {
            "_project": "PivotTable P1 handler",
            "min_action_count": 1,
            "must_include_action_types": ["pivot_table"],
        },
    },

    # ── 015 — Conditional format color scale (P1 handler) ─────────────────────
    {
        "id": "015-conditional-format-color-scale",
        "request": {
            "user_id": "test-015",
            "message": (
                "On Sheet1, apply a green-yellow-red color scale to column D "
                "(D2:D20) so high values are green and low values are red"
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:D20",
                sheet_data=[
                    ["Item", "Q1", "Q2", "Total"],
                    ["A", 100, 120, 220],
                ],
            ),
        },
        "rubric": {
            "_project": "Conditional format color scale (P1)",
            "min_action_count": 1,
            # Either the new advanced handler or the basic conditional format
            "must_include_action_types_any_of": [
                "conditional_format_advanced",
                "add_conditional_format",
            ],
        },
    },

    # ── 016 — Multi-sheet build (3 tabs from one prompt) ──────────────────────
    {
        "id": "016-multi-sheet-build",
        "request": {
            "user_id": "test-016",
            "message": (
                "Build a 3-tab workbook: Sheet 'Data' with sample sales data "
                "(month, region, sales), Sheet 'Analysis' with formulas "
                "summing sales by region, Sheet 'Summary' with a formatted "
                "header and the totals."
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="empty",
                sheet_data=[],
            ),
        },
        "rubric": {
            "_project": "Multi-sheet build pattern",
            "min_action_count": 5,
            "must_include_action_types": ["add_sheet", "write_range"],
        },
    },

    # ── 017 — Cross-sheet formula reference ───────────────────────────────────
    {
        "id": "017-cross-sheet-formula",
        "request": {
            "user_id": "test-017",
            "message": (
                "On Sheet 'Summary', cell B2, write a formula that sums "
                "all of column B on the 'Data' sheet"
            ),
            "context": _make_excel_context(
                sheet="Summary",
                all_sheets=["Summary", "Data"],
                used_range="A1:B5",
                sheet_data=[
                    ["Metric", "Value"],
                    ["Total Sales", None],
                ],
                sheet_summaries=[
                    {"name": "Summary", "used_range": "A1:B5", "rows": 5, "cols": 2,
                     "preview": [["Metric", "Value"], ["Total Sales", None]],
                     "preview_formulas": [["Metric", "Value"], ["Total Sales", None]]},
                    {"name": "Data", "used_range": "A1:B100", "rows": 100, "cols": 2,
                     "preview": [["Date", "Sales"], ["2026-01-01", 100]],
                     "preview_formulas": [["Date", "Sales"], ["2026-01-01", 100]]},
                ],
            ),
        },
        "rubric": {
            "_project": "Cross-sheet formula reference",
            "min_action_count": 1,
            "must_include_action_types": ["write_formula"],
            "formula_at_cell_contains": {
                "Summary!B2": ["Data!", "B"],
            },
        },
    },

    # ── 018 — Phantom sheet rejected (model invents a sheet) ──────────────────
    {
        "id": "018-phantom-sheet-rejected",
        "request": {
            "user_id": "test-018",
            "message": (
                "Add a summary row to the 'Transactions' tab. Bold the headers "
                "and freeze the top row."
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                # Note: NO 'Transactions' sheet exists. The model used to invent
                # this name in PlacerHills-09 sessions.
                all_sheets=["Sheet1", "Data"],
                used_range="A1:C5",
                sheet_data=[
                    ["Item", "Price", "Qty"],
                ],
            ),
        },
        "rubric": {
            "_project": "Phantom sheet — must NOT write to non-existent 'Transactions'",
            "forbidden_sheet_targets": ["Transactions"],
            # Acceptable behavior: model either asks user to clarify, or emits
            # an add_sheet for 'Transactions' explicitly. NEVER silently writes.
            "reply_not_empty": False,
        },
    },

    # ── 019 — Discuss mode: "what is a vlookup" should NOT emit actions ───────
    {
        "id": "019-discuss-mode-no-actions",
        "request": {
            "user_id": "test-019",
            "message": "what is a vlookup and when should I use it",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:A1",
                sheet_data=[],
            ),
        },
        "rubric": {
            "_project": "Discuss-mode: educational ask, no actions emitted",
            "max_action_count": 0,
            "reply_not_empty": True,
        },
    },

    # ── 020 — User asks for impossible feature, must NOT pretend ──────────────
    {
        "id": "020-impossible-feature-honest",
        "request": {
            "user_id": "test-020",
            "message": (
                "Export this workbook as a PDF and email it to me at my "
                "personal address"
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:B3",
                sheet_data=[["A", 1], ["B", 2]],
            ),
        },
        "rubric": {
            "_project": "Honest 'I cannot' for unsupported features",
            # Should NOT emit fake actions claiming success
            "reply_not_empty": True,
            # Forbidden patterns in actions: nothing that claims to email or PDF
        },
    },

    # ── 021 — Apply all suggestions follow-up (model used to invent sheets) ──
    {
        "id": "021-apply-all-changes-followup",
        "request": {
            "user_id": "test-021",
            "message": "apply all of the changes you mentioned",
            "context": _make_excel_context(
                sheet="Calculator",
                all_sheets=["Calculator", "Data"],
                used_range="A1:E10",
                sheet_data=[
                    ["Restaurant", "Value", "Buffet", "Profit", "ROI"],
                    ["A", 38.19, 11.99, 26.20, 2.18],
                    ["B", 56.39, 19.99, 36.40, 1.82],
                ],
            ),
        },
        "rubric": {
            "_project": "Apply-all follow-up: must not invent sheets",
            # Model can do various things, but must not write to sheets that
            # don't exist (the buffet 'Dashboard' bug).
            "forbidden_sheet_targets": [
                # Common phantom names from prior bug instances
                "Dashboard",
                "Summary",
                "Pivot",
            ],
        },
    },

    # ── 022 — Conditional format basic (cell-value rule) ──────────────────────
    {
        "id": "022-conditional-format-basic",
        "request": {
            "user_id": "test-022",
            "message": (
                "On Sheet1, highlight cells in column C that are greater than "
                "100 with green background"
            ),
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:C20",
                sheet_data=[
                    ["Item", "Date", "Value"],
                    ["A", "2026-01-01", 50],
                    ["B", "2026-01-02", 150],
                ],
            ),
        },
        "rubric": {
            "_project": "Basic conditional formatting",
            "min_action_count": 1,
            "must_include_action_types_any_of": [
                "add_conditional_format",
                "conditional_format_advanced",
            ],
        },
    },

    # ── 023 — Freeze panes + bold headers (polish bundle) ─────────────────────
    {
        "id": "023-freeze-and-bold-headers",
        "request": {
            "user_id": "test-023",
            "message": "Freeze the top row on Sheet1 and bold the headers",
            "context": _make_excel_context(
                sheet="Sheet1",
                all_sheets=["Sheet1"],
                used_range="A1:E50",
                sheet_data=[
                    ["Date", "Region", "Product", "Sales", "Cost"],
                ],
            ),
        },
        "rubric": {
            "_project": "Freeze + bold headers",
            "min_action_count": 1,
            "must_include_action_types_any_of": [
                "freeze_panes",
                "format_range",
            ],
        },
    },

    # ── 024 — RStudio: write Rmd + render in single run_r_code ────────────────
    {
        "id": "024-rstudio-rmd-and-render",
        "request": {
            "user_id": "test-024",
            "message": (
                "Generate an RMarkdown report named 'analysis.Rmd' that runs a "
                "linear regression on the mtcars dataset (mpg ~ wt + cyl) and "
                "knits to PDF. Use base R only."
            ),
            "context": {
                "app": "rstudio",
                "working_dir": "/Users/test/projects/r",
                "files_in_dir": [],
                "rstudio_panes": {"console": "", "source": ""},
            },
        },
        "rubric": {
            "_project": "RStudio: writeLines + rmarkdown::render in ONE run_r_code",
            "min_action_count": 1,
            "max_action_count": 2,
            "exactly_one_run_r_code": True,
            "must_include_action_types": ["run_r_code"],
            "code_contains": ["writeLines", "rmarkdown::render", "lm("],
            "reply_not_empty": False,
        },
    },
]


# ── Generator ────────────────────────────────────────────────────────────────


def write_case(spec):
    case_dir = CASES_DIR / spec["id"]
    case_dir.mkdir(parents=True, exist_ok=True)

    request_path = case_dir / "request.json"
    rubric_path = case_dir / "rubric.yaml"

    with request_path.open("w", encoding="utf-8") as f:
        json.dump(spec["request"], f, indent=2)

    rubric_with_header = {
        "_generated_by": "_build_hardening_cases.py",
        "_purpose": ("Hardening sprint regression — pin down a known good "
                     "behavior so future ships can't regress it."),
        **spec["rubric"],
    }
    with rubric_path.open("w", encoding="utf-8") as f:
        yaml.safe_dump(rubric_with_header, f, sort_keys=False, width=88)

    return case_dir


def main():
    print(f"Generating {len(CASES)} hardening cases into {CASES_DIR}/")
    for spec in CASES:
        case_dir = write_case(spec)
        print(f"  ✓ {spec['id']}")
    print(f"\nDone. Total cases now in directory:")
    for d in sorted(CASES_DIR.iterdir()):
        if d.is_dir():
            print(f"  - {d.name}")


if __name__ == "__main__":
    main()
