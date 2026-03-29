"""
Test Suite: Real-world financial analyst scenarios for tsifl Excel AI.
Sends prompts to the /chat/ endpoint and validates Claude returns correct actions.
"""

import httpx
import pytest
import json
import asyncio

BACKEND_URL = "https://focused-solace-production-6839.up.railway.app"
TEST_USER_ID = "test-excel-scenarios-001"
TIMEOUT = 60.0  # Claude can take time to respond


def make_excel_context(sheet="Sheet1", sheets_data=None, selected_cell="A1"):
    """Build a minimal Excel workbook context for testing."""
    ctx = {
        "app": "excel",
        "sheet": sheet,
        "selected_cell": selected_cell,
        "sheet_data": [],
        "sheet_formulas": [],
        "sheet_summaries": sheets_data or [],
    }
    return ctx


def make_multi_sheet_context(sheets: dict):
    """Build context with multiple sheets populated with data."""
    summaries = []
    for name, data in sheets.items():
        rows = len(data) if data else 0
        cols = max(len(r) for r in data) if data and rows > 0 else 0
        summaries.append({
            "name": name,
            "rows": rows,
            "cols": cols,
            "used_range": f"'{name}'!A1:{chr(64+max(cols,1))}{max(rows,1)}",
            "preview": data,
            "preview_formulas": [],
        })
    first_sheet = list(sheets.keys())[0]
    return make_excel_context(
        sheet=first_sheet,
        sheets_data=summaries,
        selected_cell="A1",
    )


async def send_chat(message: str, context: dict) -> dict:
    """Send a chat request to the backend and return the parsed response."""
    payload = {
        "user_id": TEST_USER_ID,
        "message": message,
        "context": context,
        "session_id": "test-session",
    }
    async with httpx.AsyncClient(timeout=TIMEOUT) as client:
        resp = await client.post(f"{BACKEND_URL}/chat/", json=payload)
        resp.raise_for_status()
        return resp.json()


def get_all_actions(result: dict) -> list:
    """Extract all actions from the response (handles single action or list)."""
    actions = result.get("actions", [])
    single = result.get("action", {})
    if single and not actions:
        actions = [single]
    return actions


def has_action_type(actions: list, action_type: str) -> bool:
    """Check if any action in the list has the given type."""
    return any(a.get("type") == action_type for a in actions)


def count_action_type(actions: list, action_type: str) -> int:
    """Count actions of a given type."""
    return sum(1 for a in actions if a.get("type") == action_type)


def actions_of_type(actions: list, action_type: str) -> list:
    """Get all actions of a given type."""
    return [a for a in actions if a.get("type") == action_type]


def any_payload_contains(actions: list, key: str, substring: str) -> bool:
    """Check if any action's payload contains a key with a value containing substring."""
    for a in actions:
        p = a.get("payload", {})
        val = str(p.get(key, ""))
        if substring.lower() in val.lower():
            return True
    return False


def sheets_targeted(actions: list) -> set:
    """Get set of sheet names targeted by actions."""
    sheets = set()
    for a in actions:
        p = a.get("payload", {})
        if "sheet" in p:
            sheets.add(p["sheet"])
    return sheets


# ─────────────────────────────────────────────────────────────────────────────
# TEST SCENARIOS
# ─────────────────────────────────────────────────────────────────────────────

class TestFinancialModeling:
    """Test financial modeling scenarios."""

    @pytest.mark.asyncio
    async def test_three_statement_model(self):
        """Build a 3-statement financial model (IS, BS, CF)."""
        ctx = make_excel_context(sheet="Sheet1")
        result = await send_chat(
            "Build a 3-statement financial model with Income Statement, Balance Sheet, "
            "and Cash Flow Statement sheets. Use sample data: Revenue $1M, COGS 60%, "
            "OpEx $200K, Tax Rate 25%, Starting Cash $500K. Link all three statements.",
            ctx
        )
        actions = get_all_actions(result)

        # Should create multiple sheets
        assert len(actions) >= 10, f"Expected 10+ actions for 3-statement model, got {len(actions)}"
        assert has_action_type(actions, "add_sheet") or has_action_type(actions, "navigate_sheet"), \
            "Should create/navigate to multiple sheets"

        # Should write formulas and cell values
        write_actions = count_action_type(actions, "write_cell") + count_action_type(actions, "write_range")
        assert write_actions >= 5, f"Expected 5+ write actions, got {write_actions}"

        # Should target multiple sheets
        targeted = sheets_targeted(actions)
        assert len(targeted) >= 2, f"Expected actions on 2+ sheets, got {targeted}"

        print(f"  ✓ 3-statement model: {len(actions)} actions across {targeted}")

    @pytest.mark.asyncio
    async def test_dcf_valuation(self):
        """DCF valuation model with WACC calculation."""
        ctx = make_excel_context(sheet="Sheet1")
        result = await send_chat(
            "Build a DCF valuation model. Assumptions: Revenue $5M growing 10% annually for 5 years, "
            "EBITDA margin 30%, CapEx 5% of revenue, D&A 3% of revenue, NWC change 2% of revenue, "
            "WACC 10%, terminal growth rate 3%. Calculate enterprise value and equity value.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 8, f"Expected 8+ actions for DCF, got {len(actions)}"

        # Should contain formulas for NPV/PV calculations
        has_pv_formula = False
        for a in actions:
            p = a.get("payload", {})
            formula = str(p.get("formula", "")) + str(p.get("value", ""))
            if any(kw in formula.upper() for kw in ["NPV", "PV", "WACC", "FCF", "TERMINAL"]):
                has_pv_formula = True
                break
        # Allow text references too
        if not has_pv_formula:
            has_pv_formula = any_payload_contains(actions, "value", "WACC") or \
                            any_payload_contains(actions, "value", "Terminal")

        assert has_pv_formula, "DCF should include PV/NPV/WACC-related formulas or labels"
        print(f"  ✓ DCF model: {len(actions)} actions")


class TestBudgetAndAnalysis:
    """Test budget and analysis scenarios."""

    @pytest.mark.asyncio
    async def test_budget_vs_actuals(self):
        """Monthly budget vs actuals variance analysis."""
        ctx = make_multi_sheet_context({
            "Budget": [
                ["Month", "Revenue", "COGS", "OpEx"],
                ["Jan", 100000, 60000, 20000],
                ["Feb", 110000, 66000, 21000],
                ["Mar", 120000, 72000, 22000],
            ],
            "Actuals": [
                ["Month", "Revenue", "COGS", "OpEx"],
                ["Jan", 105000, 62000, 19000],
                ["Feb", 108000, 67000, 22000],
                ["Mar", 125000, 70000, 23000],
            ],
        })
        result = await send_chat(
            "Create a Variance Analysis sheet that shows budget vs actuals for each month "
            "with dollar variance and percentage variance for Revenue, COGS, and OpEx. "
            "Highlight variances greater than 5% in red.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 5, f"Expected 5+ actions for variance analysis, got {len(actions)}"
        assert has_action_type(actions, "add_sheet") or has_action_type(actions, "navigate_sheet"), \
            "Should create/navigate to Variance Analysis sheet"

        # Should have formulas referencing both sheets
        has_cross_ref = False
        for a in actions:
            p = a.get("payload", {})
            f = str(p.get("formula", ""))
            if "Budget" in f or "Actuals" in f or "!" in f:
                has_cross_ref = True
                break
        # Also acceptable: write_range with calculated values
        assert has_cross_ref or count_action_type(actions, "write_cell") >= 3, \
            "Should have cross-sheet references or calculated values"

        print(f"  ✓ Budget vs actuals: {len(actions)} actions")

    @pytest.mark.asyncio
    async def test_loan_amortization(self):
        """Loan amortization schedule."""
        ctx = make_excel_context(sheet="Sheet1")
        result = await send_chat(
            "Create a loan amortization schedule for a $500,000 loan at 6% annual interest "
            "rate over 30 years with monthly payments. Show: Month, Payment, Principal, "
            "Interest, and Remaining Balance for all 360 months. Use PMT, IPMT, PPMT functions.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 5, f"Expected 5+ actions for amortization, got {len(actions)}"

        # Should use fill_down (NOT 360 individual write_cell actions)
        assert has_action_type(actions, "fill_down"), \
            "Must use fill_down for repeating formulas, not individual write_cell for each row"

        # Should NOT have excessive write_cell actions for the formula rows
        write_cells = count_action_type(actions, "write_cell")
        assert write_cells < 30, \
            f"Should use fill_down instead of {write_cells} individual write_cell actions"

        # Should contain PMT-related formulas
        has_pmt = False
        for a in actions:
            f = str(a.get("payload", {}).get("formula", "")).upper()
            if any(fn in f for fn in ["PMT", "IPMT", "PPMT"]):
                has_pmt = True
                break
        assert has_pmt, "Should use PMT/IPMT/PPMT Excel functions"

        print(f"  ✓ Loan amortization: {len(actions)} actions, uses fill_down correctly")


class TestPortfolioAndData:
    """Test portfolio and data analysis scenarios."""

    @pytest.mark.asyncio
    async def test_portfolio_tracker(self):
        """Portfolio performance tracker with returns calculation."""
        ctx = make_multi_sheet_context({
            "Portfolio": [
                ["Ticker", "Shares", "Buy Price", "Current Price"],
                ["AAPL", 100, 150.00, 185.00],
                ["GOOGL", 50, 2800.00, 2950.00],
                ["MSFT", 75, 300.00, 380.00],
                ["AMZN", 30, 3200.00, 3450.00],
                ["TSLA", 60, 250.00, 220.00],
            ],
        })
        result = await send_chat(
            "Add columns for: Market Value, Cost Basis, Gain/Loss ($), Gain/Loss (%), "
            "Portfolio Weight. Add a totals row at bottom. Format currency as $#,##0.00 "
            "and percentages as 0.00%.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 5, f"Expected 5+ actions for portfolio tracker, got {len(actions)}"

        # Should have formulas for calculations
        has_formula = any(
            a.get("payload", {}).get("formula", "")
            for a in actions
        )
        assert has_formula or has_action_type(actions, "fill_down"), \
            "Should include formulas for calculations"

        # Should have number formatting
        assert has_action_type(actions, "set_number_format") or has_action_type(actions, "format_range"), \
            "Should format numbers (currency/percentage)"

        print(f"  ✓ Portfolio tracker: {len(actions)} actions")

    @pytest.mark.asyncio
    async def test_pivot_summary(self):
        """Pivot-style summary from raw transaction data."""
        ctx = make_multi_sheet_context({
            "Transactions": [
                ["Date", "Region", "Product", "Amount"],
                ["2024-01-01", "North", "Widget A", 1500],
                ["2024-01-02", "South", "Widget B", 2300],
                ["2024-01-03", "North", "Widget A", 1800],
                ["2024-01-04", "East", "Widget C", 900],
                ["2024-01-05", "South", "Widget A", 2100],
                ["2024-01-06", "North", "Widget B", 1700],
                ["2024-01-07", "East", "Widget A", 1400],
                ["2024-01-08", "West", "Widget C", 2500],
                ["2024-01-09", "North", "Widget C", 1100],
                ["2024-01-10", "South", "Widget B", 1950],
            ],
        })
        result = await send_chat(
            "Create a Summary sheet with a pivot-style breakdown: "
            "Revenue by Region and Revenue by Product using SUMIFS. "
            "Also show grand total, average transaction, and count of transactions.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 5, f"Expected 5+ actions for pivot summary, got {len(actions)}"

        # Should use SUMIFS or SUMIF
        has_sumifs = False
        for a in actions:
            f = str(a.get("payload", {}).get("formula", "")).upper()
            if "SUMIF" in f:
                has_sumifs = True
                break
        assert has_sumifs, "Should use SUMIFS for pivot-style aggregation"

        # Should create/navigate to Summary sheet
        assert has_action_type(actions, "add_sheet") or \
               any_payload_contains(actions, "sheet", "Summary"), \
            "Should create or target a Summary sheet"

        print(f"  ✓ Pivot summary: {len(actions)} actions with SUMIFS")


class TestFormulasAndReferences:
    """Test formula and cross-sheet reference scenarios."""

    @pytest.mark.asyncio
    async def test_vlookup_index_match(self):
        """VLOOKUP/INDEX-MATCH equivalent formulas across sheets."""
        ctx = make_multi_sheet_context({
            "Products": [
                ["Product ID", "Product Name", "Category", "Unit Price"],
                ["P001", "Laptop", "Electronics", 999.99],
                ["P002", "Mouse", "Accessories", 29.99],
                ["P003", "Monitor", "Electronics", 499.99],
                ["P004", "Keyboard", "Accessories", 79.99],
            ],
            "Orders": [
                ["Order ID", "Product ID", "Quantity"],
                ["ORD001", "P001", 5],
                ["ORD002", "P003", 10],
                ["ORD003", "P002", 25],
                ["ORD004", "P004", 15],
            ],
        })
        result = await send_chat(
            "In the Orders sheet, add columns D (Product Name), E (Unit Price), "
            "and F (Line Total) using INDEX-MATCH to look up from the Products sheet. "
            "Use fill_down for the formula rows.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 3, f"Expected 3+ actions for INDEX-MATCH, got {len(actions)}"

        # Should use INDEX/MATCH or VLOOKUP
        has_lookup = False
        for a in actions:
            f = str(a.get("payload", {}).get("formula", "")).upper()
            if any(fn in f for fn in ["INDEX", "MATCH", "VLOOKUP", "XLOOKUP"]):
                has_lookup = True
                break
        assert has_lookup, "Should use INDEX-MATCH or VLOOKUP/XLOOKUP"

        # Should use fill_down
        assert has_action_type(actions, "fill_down"), "Should use fill_down for formula columns"

        print(f"  ✓ INDEX-MATCH across sheets: {len(actions)} actions")

    @pytest.mark.asyncio
    async def test_multi_sheet_workbook_setup(self):
        """Multi-sheet workbook setup with cross-sheet references."""
        ctx = make_excel_context(sheet="Sheet1")
        result = await send_chat(
            "Set up a workbook with 4 sheets: Assumptions, Revenue, Expenses, Dashboard. "
            "In Assumptions, put growth rate (10%), tax rate (25%), discount rate (8%). "
            "In Revenue, create Year 1-5 projections starting at $1M with growth from Assumptions. "
            "In Expenses, create Year 1-5 at 60% of Revenue. "
            "In Dashboard, summarize net income (Revenue - Expenses) for each year.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 10, f"Expected 10+ actions for multi-sheet setup, got {len(actions)}"

        # Should create 3-4 new sheets
        add_sheets = count_action_type(actions, "add_sheet")
        nav_sheets = count_action_type(actions, "navigate_sheet")
        assert add_sheets >= 2 or nav_sheets >= 2, \
            f"Should create/navigate multiple sheets (add: {add_sheets}, nav: {nav_sheets})"

        # Should target 3+ different sheets
        targeted = sheets_targeted(actions)
        assert len(targeted) >= 3, f"Expected 3+ sheets targeted, got {targeted}"

        print(f"  ✓ Multi-sheet setup: {len(actions)} actions across {targeted}")


class TestFormattingAndCharts:
    """Test conditional formatting, charts, and data validation."""

    @pytest.mark.asyncio
    async def test_conditional_formatting(self):
        """Conditional formatting rules (highlight negatives, top 10%)."""
        ctx = make_multi_sheet_context({
            "P&L": [
                ["", "Q1", "Q2", "Q3", "Q4"],
                ["Revenue", 500000, 520000, 480000, 550000],
                ["COGS", -300000, -310000, -290000, -330000],
                ["Gross Profit", 200000, 210000, 190000, 220000],
                ["OpEx", -120000, -125000, -130000, -118000],
                ["Net Income", 80000, 85000, 60000, 102000],
            ],
        })
        result = await send_chat(
            "Format the P&L sheet: Make headers bold with blue (#0D5EAF) background and white text. "
            "Format all numbers as $#,##0. Make negative numbers red. "
            "Bold the Net Income row. Add borders to the entire table.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 3, f"Expected 3+ formatting actions, got {len(actions)}"

        # Should have format_range or set_number_format
        format_count = count_action_type(actions, "format_range") + \
                      count_action_type(actions, "set_number_format")
        assert format_count >= 1, "Should include formatting actions"

        # Check for bold formatting
        has_bold = False
        for a in actions:
            p = a.get("payload", {})
            if p.get("bold") is True:
                has_bold = True
                break
        assert has_bold, "Should apply bold formatting"

        print(f"  ✓ Conditional formatting: {len(actions)} actions, {format_count} format ops")

    @pytest.mark.asyncio
    async def test_chart_creation(self):
        """Chart creation (bar, line, combo charts)."""
        ctx = make_multi_sheet_context({
            "Sales": [
                ["Month", "Revenue", "Expenses", "Profit"],
                ["Jan", 100000, 70000, 30000],
                ["Feb", 110000, 75000, 35000],
                ["Mar", 120000, 72000, 48000],
                ["Apr", 115000, 78000, 37000],
                ["May", 130000, 80000, 50000],
                ["Jun", 140000, 85000, 55000],
            ],
        })
        result = await send_chat(
            "Create a combo chart showing Revenue and Expenses as bars and Profit as a line. "
            "Title it 'Monthly P&L Overview'. Put it on a new Charts sheet.",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 1, "Should return at least 1 action"

        # Must use add_chart, NOT run_shell_command
        assert has_action_type(actions, "add_chart"), \
            f"Should use add_chart action, not run_shell_command. Got types: {[a.get('type') for a in actions]}"
        assert not has_action_type(actions, "run_shell_command"), \
            "Should NOT use run_shell_command for chart creation — use add_chart instead"

        action_types = [a.get("type") for a in actions]
        print(f"  ✓ Chart creation: {len(actions)} actions, types: {action_types}")

    @pytest.mark.asyncio
    async def test_data_validation_dropdowns(self):
        """Data validation dropdowns."""
        ctx = make_multi_sheet_context({
            "Data Entry": [
                ["Employee", "Department", "Status", "Amount"],
                ["", "", "", ""],
            ],
            "Lists": [
                ["Departments", "Statuses"],
                ["Engineering", "Active"],
                ["Sales", "Inactive"],
                ["Marketing", "On Leave"],
                ["Finance", "Terminated"],
            ],
        })
        result = await send_chat(
            "In the Data Entry sheet, set up data validation: "
            "Column B (Department) should have a dropdown with values from Lists!A2:A5. "
            "Column C (Status) should have a dropdown with values from Lists!B2:B5.",
            ctx
        )
        actions = get_all_actions(result)
        assert len(actions) >= 1, "Should return at least 1 action"

        # Must use add_data_validation, NOT run_shell_command
        assert has_action_type(actions, "add_data_validation"), \
            f"Should use add_data_validation action, not run_shell_command. Got types: {[a.get('type') for a in actions]}"
        assert not has_action_type(actions, "run_shell_command"), \
            "Should NOT use run_shell_command for data validation — use add_data_validation instead"

        action_types = [a.get("type") for a in actions]
        print(f"  ✓ Data validation: {len(actions)} actions, types: {action_types}")


class TestImportAndAnalysis:
    """Test the import CSV + build analysis flow."""

    @pytest.mark.asyncio
    async def test_import_csv_and_analyze(self):
        """Import CSV then build full analysis (the existing flow)."""
        ctx = make_excel_context(sheet="Sheet1")
        result = await send_chat(
            "Import /tmp/sales_data.csv into Excel and then create an Analysis sheet with: "
            "1. Total Revenue using SUM "
            "2. Average Unit Price "
            "3. Revenue by Region using SUMIFS "
            "4. Top selling product using INDEX-MATCH with MAX",
            ctx
        )
        actions = get_all_actions(result)

        assert len(actions) >= 3, f"Expected 3+ actions for import+analyze, got {len(actions)}"

        # Should import CSV exactly once
        import_count = count_action_type(actions, "import_csv")
        assert import_count == 1, f"Should import CSV exactly once, got {import_count}"

        # Should create analysis sheet
        assert has_action_type(actions, "add_sheet") or has_action_type(actions, "navigate_sheet"), \
            "Should create/navigate to Analysis sheet"

        # Should have analysis formulas
        has_analysis = False
        for a in actions:
            f = str(a.get("payload", {}).get("formula", "")).upper()
            if any(fn in f for fn in ["SUM", "AVERAGE", "SUMIFS", "INDEX", "MATCH"]):
                has_analysis = True
                break
        assert has_analysis, "Should include analysis formulas (SUM, AVERAGE, SUMIFS, etc.)"

        print(f"  ✓ Import + analyze: {len(actions)} actions, 1 import")


# ─────────────────────────────────────────────────────────────────────────────
# RUNNER
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    """Run all tests and print a summary."""
    import sys

    async def run_all():
        results = []
        scenarios = [
            ("3-Statement Financial Model", TestFinancialModeling().test_three_statement_model),
            ("DCF Valuation", TestFinancialModeling().test_dcf_valuation),
            ("Budget vs Actuals", TestBudgetAndAnalysis().test_budget_vs_actuals),
            ("Loan Amortization", TestBudgetAndAnalysis().test_loan_amortization),
            ("Portfolio Tracker", TestPortfolioAndData().test_portfolio_tracker),
            ("Pivot Summary", TestPortfolioAndData().test_pivot_summary),
            ("VLOOKUP/INDEX-MATCH", TestFormulasAndReferences().test_vlookup_index_match),
            ("Multi-Sheet Workbook", TestFormulasAndReferences().test_multi_sheet_workbook_setup),
            ("Conditional Formatting", TestFormattingAndCharts().test_conditional_formatting),
            ("Chart Creation", TestFormattingAndCharts().test_chart_creation),
            ("Data Validation", TestFormattingAndCharts().test_data_validation_dropdowns),
            ("Import CSV + Analyze", TestImportAndAnalysis().test_import_csv_and_analyze),
        ]

        for name, test_fn in scenarios:
            try:
                await test_fn()
                results.append((name, "PASS", ""))
            except AssertionError as e:
                results.append((name, "FAIL", str(e)))
            except Exception as e:
                results.append((name, "ERROR", str(e)))

        # Print summary
        print("\n" + "="*70)
        print("TEST RESULTS SUMMARY")
        print("="*70)
        passed = sum(1 for _, s, _ in results if s == "PASS")
        failed = sum(1 for _, s, _ in results if s == "FAIL")
        errors = sum(1 for _, s, _ in results if s == "ERROR")
        for name, status, msg in results:
            icon = "✅" if status == "PASS" else "❌" if status == "FAIL" else "⚠️"
            line = f"{icon} {name}: {status}"
            if msg:
                line += f" — {msg[:100]}"
            print(line)
        print(f"\nTotal: {len(results)} | Passed: {passed} | Failed: {failed} | Errors: {errors}")
        print(f"Pass rate: {passed/len(results)*100:.0f}%")
        return results

    results = asyncio.run(run_all())
    sys.exit(0 if all(s == "PASS" for _, s, _ in results) else 1)
