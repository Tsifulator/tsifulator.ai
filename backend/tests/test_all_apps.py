"""
Test Suite: Multi-app scenarios for tsifl AI across all supported platforms.
Sends prompts to the /chat/ endpoint and validates Claude returns correct actions.

Tests 5+ scenarios per app context:
- Excel (existing, expanded)
- PowerPoint
- Word
- Gmail
- VS Code
- Google Sheets
- Google Docs
- Google Slides
- Browser
"""

import httpx
import pytest
import json
import asyncio

BACKEND_URL = "https://focused-solace-production-6839.up.railway.app"
TEST_USER_ID = "test-all-apps-001"
TIMEOUT = 60.0


async def send_chat(message: str, context: dict) -> dict:
    payload = {
        "user_id": TEST_USER_ID,
        "message": message,
        "context": context,
        "session_id": "test-" + context.get("app", "unknown"),
    }
    async with httpx.AsyncClient(timeout=TIMEOUT) as client:
        resp = await client.post(f"{BACKEND_URL}/chat/", json=payload)
        assert resp.status_code == 200, f"Backend returned {resp.status_code}: {resp.text}"
        return resp.json()


def get_actions(result):
    actions = result.get("actions", [])
    if not actions and result.get("action", {}).get("type"):
        actions = [result["action"]]
    return actions


def action_types(actions):
    return [a["type"] for a in actions]


# ── PowerPoint Tests ────────────────────────────────────────────────────────

PPT_CONTEXT = {
    "app": "powerpoint",
    "total_slides": 0,
    "slides": [],
}

PPT_CONTEXT_WITH_SLIDES = {
    "app": "powerpoint",
    "total_slides": 3,
    "current_slide": {"index": 0, "layout": "Title Slide"},
    "slides": [
        {"index": 0, "title": "Q4 Results", "shapes": [{"type": "TextBox", "text": "Q4 Results"}]},
        {"index": 1, "title": "Revenue", "shapes": [{"type": "TextBox", "text": "Revenue Analysis"}]},
        {"index": 2, "title": "Outlook", "shapes": [{"type": "TextBox", "text": "2025 Outlook"}]},
    ],
}


@pytest.mark.asyncio
async def test_ppt_create_title_slide():
    result = await send_chat("Create a title slide for 'Q4 2025 Board Meeting'", PPT_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 1
    types = action_types(actions)
    assert "create_slide" in types or "add_text_box" in types


@pytest.mark.asyncio
async def test_ppt_create_pitch_deck():
    result = await send_chat("Create a 5-slide pitch deck: Title, Problem, Solution, Market, Ask", PPT_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 5
    create_count = sum(1 for a in actions if a["type"] == "create_slide")
    assert create_count >= 5


@pytest.mark.asyncio
async def test_ppt_add_table():
    result = await send_chat("Add a table with quarterly revenue: Q1=$10M, Q2=$12M, Q3=$15M, Q4=$18M", PPT_CONTEXT_WITH_SLIDES)
    actions = get_actions(result)
    types = action_types(actions)
    assert "add_table" in types or "add_text_box" in types


@pytest.mark.asyncio
async def test_ppt_add_chart():
    result = await send_chat("Add a bar chart showing revenue by quarter", PPT_CONTEXT_WITH_SLIDES)
    actions = get_actions(result)
    assert len(actions) >= 1


@pytest.mark.asyncio
async def test_ppt_no_shell_command():
    result = await send_chat("Add a background color to slide 1", PPT_CONTEXT_WITH_SLIDES)
    actions = get_actions(result)
    types = action_types(actions)
    assert "run_shell_command" not in types


# ── Word Tests ──────────────────────────────────────────────────────────────

WORD_CONTEXT = {
    "app": "word",
    "total_paragraphs": 0,
    "paragraphs": [],
    "tables": [],
}

WORD_CONTEXT_WITH_CONTENT = {
    "app": "word",
    "total_paragraphs": 10,
    "paragraphs": [
        {"text": "Investment Memo", "style": "Heading 1"},
        {"text": "Executive Summary", "style": "Heading 2"},
        {"text": "The company has shown strong growth...", "style": "Normal"},
    ],
    "tables": [],
    "selection": "",
}


@pytest.mark.asyncio
async def test_word_create_memo():
    result = await send_chat("Write a financial memo about Q4 results", WORD_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 3
    types = action_types(actions)
    assert any(t in types for t in ["insert_paragraph", "insert_text"])


@pytest.mark.asyncio
async def test_word_insert_table():
    result = await send_chat("Insert a table with 3 columns: Metric, Q3, Q4 — with Revenue $10M/$12M and EBITDA $3M/$4M", WORD_CONTEXT_WITH_CONTENT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "insert_table" in types


@pytest.mark.asyncio
async def test_word_format_heading():
    result = await send_chat("Add a heading 'Financial Analysis' and two paragraphs about revenue trends", WORD_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 2


@pytest.mark.asyncio
async def test_word_find_replace():
    result = await send_chat("Replace all instances of 'company' with 'Acme Corp'", WORD_CONTEXT_WITH_CONTENT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "find_and_replace" in types


@pytest.mark.asyncio
async def test_word_no_shell_command():
    result = await send_chat("Set page margins to 1 inch all sides", WORD_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "run_shell_command" not in types


# ── Gmail Tests ─────────────────────────────────────────────────────────────

GMAIL_CONTEXT = {
    "app": "gmail",
    "email": "test@example.com",
    "recent_emails": [
        {"from": "john@acme.com", "subject": "Q4 Board Materials"},
        {"from": "sarah@investor.com", "subject": "Follow up on term sheet"},
    ],
}

GMAIL_THREAD_CONTEXT = {
    "app": "gmail",
    "email": "test@example.com",
    "current_thread": {
        "subject": "Partnership Proposal",
        "messages": [
            {"from": "partner@bigco.com", "snippet": "We'd like to discuss a potential partnership. Our team can meet next week."},
        ],
    },
}


@pytest.mark.asyncio
async def test_gmail_draft_email():
    result = await send_chat("Draft a follow-up email to sarah@investor.com about the term sheet", GMAIL_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "draft_email" in types or "send_email" in types


@pytest.mark.asyncio
async def test_gmail_reply():
    result = await send_chat("Reply to this thread accepting the meeting — suggest Tuesday at 2pm", GMAIL_THREAD_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 1


@pytest.mark.asyncio
async def test_gmail_summarize():
    result = await send_chat("Summarize this email thread", GMAIL_THREAD_CONTEXT)
    assert result.get("reply"), "Should have a text reply"


@pytest.mark.asyncio
async def test_gmail_action_items():
    result = await send_chat("Extract action items from this thread", GMAIL_THREAD_CONTEXT)
    assert result.get("reply"), "Should have a text reply"


@pytest.mark.asyncio
async def test_gmail_cold_outreach():
    result = await send_chat("Draft a cold outreach email to ceo@startup.com introducing our analytics platform", GMAIL_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "draft_email" in types or "send_email" in types


# ── VS Code Tests ───────────────────────────────────────────────────────────

VSCODE_CONTEXT = {
    "app": "vscode",
    "workspace": "my-project",
    "current_file": "/Users/dev/my-project/src/utils.py",
    "language": "python",
    "line_count": 50,
    "selection": "def calculate_total(items):\n    total = 0\n    for item in items:\n        total += item.price\n    return total",
    "diagnostics": [],
    "git_branch": "main",
    "git_changes": 2,
}

VSCODE_ERROR_CONTEXT = {
    "app": "vscode",
    "workspace": "my-project",
    "current_file": "/Users/dev/my-project/src/api.py",
    "language": "python",
    "line_count": 100,
    "file_content": "import requests\n\ndef fetch_data(url):\n    resp = requests.get(url)\n    data = resp.json\n    return data['results']\n",
    "diagnostics": [
        {"file": "api.py", "line": 5, "severity": "error", "message": "'method' object is not subscriptable"},
    ],
}


@pytest.mark.asyncio
async def test_vscode_explain_code():
    result = await send_chat("Explain this code", VSCODE_CONTEXT)
    assert result.get("reply"), "Should explain the selected code"


@pytest.mark.asyncio
async def test_vscode_refactor():
    result = await send_chat("Refactor this to use a list comprehension", VSCODE_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert any(t in types for t in ["replace_selection", "insert_code"])


@pytest.mark.asyncio
async def test_vscode_fix_error():
    result = await send_chat("Fix the error in this file", VSCODE_ERROR_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert any(t in types for t in ["replace_selection", "edit_file", "insert_code"])


@pytest.mark.asyncio
async def test_vscode_generate_tests():
    result = await send_chat("Generate pytest tests for calculate_total", VSCODE_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "create_file" in types


@pytest.mark.asyncio
async def test_vscode_create_file():
    result = await send_chat("Create a new file src/models.py with a User dataclass with name, email, role fields", VSCODE_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "create_file" in types


# ── Google Sheets Tests ─────────────────────────────────────────────────────

GSHEETS_CONTEXT = {
    "app": "google_sheets",
    "spreadsheet_name": "Q4 Budget",
    "sheet_name": "Revenue",
    "all_sheets": ["Revenue", "Expenses", "Summary"],
    "active_cell": "A1",
    "data_range": "A1:D10",
    "row_count": 10,
    "col_count": 4,
    "data": [
        ["Product", "Q1", "Q2", "Q3"],
        ["Widget A", 100000, 120000, 135000],
        ["Widget B", 80000, 95000, 110000],
        ["Widget C", 50000, 60000, 70000],
    ],
    "formulas": [[], [], [], []],
}


@pytest.mark.asyncio
async def test_gsheets_add_formula():
    result = await send_chat("Add a Total row summing each quarter column", GSHEETS_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 1
    types = action_types(actions)
    assert "write_cell" in types or "write_range" in types


@pytest.mark.asyncio
async def test_gsheets_format():
    result = await send_chat("Format the header row bold with blue background", GSHEETS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "format_range" in types


@pytest.mark.asyncio
async def test_gsheets_add_chart():
    result = await send_chat("Create a bar chart of revenue by product", GSHEETS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "add_chart" in types


@pytest.mark.asyncio
async def test_gsheets_sort():
    result = await send_chat("Sort the data by Q3 revenue descending", GSHEETS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "sort_range" in types


@pytest.mark.asyncio
async def test_gsheets_add_sheet():
    result = await send_chat("Create a new sheet called 'Analysis' with a summary of total revenue per quarter", GSHEETS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "add_sheet" in types or "navigate_sheet" in types


# ── Google Docs Tests ───────────────────────────────────────────────────────

GDOCS_CONTEXT = {
    "app": "google_docs",
    "document_name": "Investment Memo",
    "paragraph_count": 5,
    "paragraphs": [
        {"text": "Investment Memo", "heading": "TITLE", "alignment": "CENTER"},
        {"text": "Company Overview", "heading": "HEADING1", "alignment": "LEFT"},
        {"text": "Acme Corp is a SaaS company...", "heading": "NORMAL", "alignment": "LEFT"},
    ],
}


@pytest.mark.asyncio
async def test_gdocs_add_section():
    result = await send_chat("Add a 'Financial Analysis' section with a paragraph about revenue growth", GDOCS_CONTEXT)
    actions = get_actions(result)
    assert len(actions) >= 2


@pytest.mark.asyncio
async def test_gdocs_insert_table():
    result = await send_chat("Insert a table with financial metrics: Revenue $50M, EBITDA $15M, Net Income $8M", GDOCS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "insert_table" in types


@pytest.mark.asyncio
async def test_gdocs_find_replace():
    result = await send_chat("Replace 'Acme Corp' with 'Acme Inc.' throughout the document", GDOCS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "find_and_replace" in types


@pytest.mark.asyncio
async def test_gdocs_add_header():
    result = await send_chat("Add a header with 'CONFIDENTIAL' and a footer with page numbers", GDOCS_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "insert_header" in types


@pytest.mark.asyncio
async def test_gdocs_write_memo():
    result = await send_chat("Write a professional memo about quarterly results", {"app": "google_docs", "document_name": "New Doc", "paragraph_count": 0, "paragraphs": []})
    actions = get_actions(result)
    assert len(actions) >= 3


# ── Google Slides Tests ─────────────────────────────────────────────────────

GSLIDES_CONTEXT = {
    "app": "google_slides",
    "presentation_name": "Board Deck",
    "slide_count": 2,
    "current_slide_index": 0,
    "slides": [
        {"index": 0, "id": "s1", "title": "Board Meeting", "shapes": [{"type": "RECTANGLE", "text": "Board Meeting Q4"}]},
        {"index": 1, "id": "s2", "title": "Agenda", "shapes": [{"type": "TEXT_BOX", "text": "1. Financials\n2. Strategy"}]},
    ],
}


@pytest.mark.asyncio
async def test_gslides_create_slide():
    result = await send_chat("Add a new slide with title 'Financial Performance' and a table of revenue by quarter", GSLIDES_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "create_slide" in types or "add_text_box" in types


@pytest.mark.asyncio
async def test_gslides_add_shapes():
    result = await send_chat("Add a blue rectangle with text 'Revenue: $50M' on slide 1", GSLIDES_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "add_shape" in types or "add_text_box" in types


@pytest.mark.asyncio
async def test_gslides_add_table():
    result = await send_chat("Add a table to slide 2: Metric / Q3 / Q4 — Revenue $40M/$50M, EBITDA $12M/$15M", GSLIDES_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "add_table" in types


@pytest.mark.asyncio
async def test_gslides_background():
    result = await send_chat("Set the background of slide 1 to dark blue", GSLIDES_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "set_slide_background" in types


@pytest.mark.asyncio
async def test_gslides_delete():
    result = await send_chat("Delete slide 2", GSLIDES_CONTEXT)
    actions = get_actions(result)
    types = action_types(actions)
    assert "delete_slide" in types


# ── Browser Tests ───────────────────────────────────────────────────────────

BROWSER_CONTEXT = {
    "app": "browser",
    "url": "https://example.com/article",
    "title": "Top 10 SaaS Metrics for 2025",
    "page_text": "The most important SaaS metrics are ARR, MRR, churn rate, CAC, LTV, net revenue retention, gross margin, burn rate, runway, and magic number. ARR (Annual Recurring Revenue) represents the annualized value of recurring subscription revenue. MRR is the monthly equivalent.",
}


@pytest.mark.asyncio
async def test_browser_summarize():
    result = await send_chat("Summarize this page", BROWSER_CONTEXT)
    # Browser context: Claude may put summary in reply or in action payload
    reply = result.get("reply", "")
    actions = get_actions(result)
    has_content = len(reply) > 10 or len(actions) > 0
    assert has_content, "Should have a text summary or actions"


@pytest.mark.asyncio
async def test_browser_extract():
    result = await send_chat("Extract the list of SaaS metrics from this page", BROWSER_CONTEXT)
    # Response could be in reply or actions (tool_choice forces tool call)
    assert result.get("reply") or get_actions(result), "Should extract the metrics"


@pytest.mark.asyncio
async def test_browser_explain():
    result = await send_chat("What is ARR and how is it different from MRR?", BROWSER_CONTEXT)
    assert result.get("reply") or get_actions(result), "Should explain the concepts"


@pytest.mark.asyncio
async def test_browser_with_selection():
    ctx = {**BROWSER_CONTEXT, "selection": "ARR (Annual Recurring Revenue) represents the annualized value of recurring subscription revenue."}
    result = await send_chat("Explain this in simpler terms", ctx)
    assert result.get("reply") or get_actions(result), "Should explain the selection"


@pytest.mark.asyncio
async def test_browser_action_items():
    ctx = {
        "app": "browser",
        "url": "https://example.com/meeting-notes",
        "title": "Team Meeting Notes",
        "page_text": "Action items: John to prepare budget by Friday. Sarah to review term sheet. Mike to schedule investor call next week.",
    }
    result = await send_chat("Extract action items with owners and deadlines", ctx)
    assert result.get("reply") or get_actions(result), "Should extract action items"
