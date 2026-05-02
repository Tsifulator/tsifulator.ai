"""
tsifl Template Engine
Generates IB-grade Excel and PowerPoint files server-side.
No Office.js limitations. No locale bugs. Deterministic output every time.

Supported templates:
  - Trading Comps (Excel + PPT)
  - Precedent Transactions (Excel)  [coming]
  - DCF (Excel)                     [coming]
"""

import io
from typing import Optional
from statistics import median, mean

# ── Shared brand constants ─────────────────────────────────────────────────
NAVY       = "002366"
NAVY_MED   = "1F3864"
BLUE_LIGHT = "D6E4F0"
BLUE_FILL  = "F0F5FA"   # very subtle alternating — less aggressive
WHITE      = "FFFFFF"
DARK_TEXT  = "1A1A2E"
GREY_LINE  = "BDC3C7"
GREY_MED   = "D1D5DB"
GREEN_CHK  = "27AE60"
RED_NEG    = "C0392B"
FONT_NAME  = "Calibri"


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL — Trading Comps
# ══════════════════════════════════════════════════════════════════════════════

def generate_comp_table_xlsx(payload: dict) -> bytes:
    """
    Generate a fully-formatted IB trading comps table as .xlsx bytes.
    Sorted by market cap descending. Includes Median, Mean, High, Low.
    """
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Trading Comps"

    companies = payload.get("companies", [])
    title     = payload.get("title", "Trading Comps")
    date_str  = payload.get("date", "")
    currency  = payload.get("currency", "USD ($M)")

    # Sort by market cap descending (IB standard)
    companies = sorted(companies, key=lambda c: c.get("market_cap") or 0, reverse=True)

    # ── Style helpers ─────────────────────────────────────────────────────
    def _fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def _font(bold=False, color=DARK_TEXT, size=10, italic=False):
        return Font(bold=bold, color=color, size=size, name=FONT_NAME, italic=italic)

    def _align(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    thin_side  = Side(style="thin", color=GREY_LINE)
    thick_side = Side(style="medium", color=NAVY)
    hair_side  = Side(style="hair", color=GREY_MED)

    def _all_border():
        return Border(top=hair_side, bottom=hair_side, left=hair_side, right=hair_side)

    def _header_border():
        return Border(
            top=thick_side, bottom=thick_side,
            left=Side(style="thin", color=NAVY_MED),
            right=Side(style="thin", color=NAVY_MED),
        )

    def _stat_border(is_last=False):
        return Border(
            top=Side(style="thin", color=NAVY),
            bottom=Side(style="medium" if is_last else "thin", color=NAVY),
            left=hair_side, right=hair_side,
        )

    # ── Column layout (IB standard order) ──────────────────────────────────
    # A=Company, B=Ticker, C=Price, D=Mkt Cap, E=Net Debt, F=EV,
    # G=Revenue, H=EBITDA, I=Gross%, J=EBITDA%, K=EV/Rev, L=EV/EBITDA, M=P/E
    cols = ["A","B","C","D","E","F","G","H","I","J","K","L","M"]
    col_widths = [24, 8, 10, 12, 12, 12, 12, 12, 11, 11, 10, 10, 10]
    for col, w in zip(cols, col_widths):
        ws.column_dimensions[col].width = w

    # ── Row 1 — Title bar ──────────────────────────────────────────────────
    ws.merge_cells("A1:M1")
    ws.row_dimensions[1].height = 30
    c = ws["A1"]
    c.value     = f"  {title}"
    c.font      = _font(bold=True, color=WHITE, size=13)
    c.fill      = _fill(NAVY)
    c.alignment = _align("left", "center")
    c.border    = Border(bottom=thick_side)

    # ── Row 2 — Subtitle ──────────────────────────────────────────────────
    ws.merge_cells("A2:D2")
    ws.row_dimensions[2].height = 18
    ws["A2"].value     = f"  {date_str}"
    ws["A2"].font      = _font(italic=True, color="6B7280", size=9)
    ws["A2"].alignment = _align("left", "center")

    ws.merge_cells("E2:M2")
    ws["E2"].value     = f"All figures in {currency}  "
    ws["E2"].font      = _font(italic=True, color="6B7280", size=9)
    ws["E2"].alignment = _align("right", "center")

    # ── Row 3 — Spacer ────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 4

    # ── Row 4 — Section sub-headers ───────────────────────────────────────
    ws.row_dimensions[4].height = 15
    section_defs = [
        ("A4:B4", ""),
        ("C4:F4", "Market Data"),
        ("G4:J4", "Financials (LTM)"),
        ("K4:M4", "Multiples"),
    ]
    for rng, label in section_defs:
        ws.merge_cells(rng)
        cell = ws[rng.split(":")[0]]
        cell.value     = label
        cell.font      = _font(bold=True, color=WHITE, size=8)
        cell.fill      = _fill(NAVY_MED)
        cell.alignment = _align("center", "center")
        cell.border    = Border(bottom=Side(style="thin", color=NAVY))

    # ── Row 5 — Column headers ────────────────────────────────────────────
    ws.row_dimensions[5].height = 32
    headers = [
        "Company", "Ticker", "Share\nPrice ($)",
        "Market Cap\n($B)", "Net Debt\n($B)", "Enterprise\nValue ($B)",
        "Revenue\n($M)", "EBITDA\n($M)", "Gross\nMargin", "EBITDA\nMargin",
        "EV /\nRevenue", "EV /\nEBITDA", "P / E"
    ]
    for col_letter, header_text in zip(cols, headers):
        c = ws[f"{col_letter}5"]
        c.value     = header_text
        c.font      = _font(bold=True, color=WHITE, size=9)
        c.fill      = _fill(NAVY)
        c.alignment = _align("center", "center", wrap=True)
        c.border    = _header_border()

    # ── Data rows ──────────────────────────────────────────────────────────
    FMT_PRICE  = '[$-409]#,##0.00'
    FMT_DEC1   = '#,##0.0'
    FMT_DEC0   = '#,##0'
    FMT_MULT   = '0.0"x"'
    FMT_PCT    = '0.0%'

    data_start = 6

    for i, co in enumerate(companies):
        row = data_start + i
        ws.row_dimensions[row].height = 20
        row_fill = _fill(BLUE_FILL) if i % 2 == 1 else _fill(WHITE)

        def _cell(col_letter, value, fmt=None, h_align="right", bold=False, neg_red=False):
            c = ws[f"{col_letter}{row}"]
            c.value     = value
            c.fill      = row_fill
            c.alignment = _align(h_align, "center")
            c.border    = _all_border()
            if fmt:
                c.number_format = fmt
            # Color: negative numbers in red
            text_color = DARK_TEXT
            if neg_red and value is not None:
                try:
                    if float(value) < 0:
                        text_color = RED_NEG
                except (ValueError, TypeError):
                    pass
            c.font = _font(bold=bold, size=9, color=text_color)

        # Company + Ticker
        _cell("A", co.get("name", ""),     h_align="left")
        _cell("B", co.get("ticker", ""),   h_align="center", bold=True)

        # Market data
        _cell("C", co.get("share_price"),  FMT_PRICE)
        _cell("D", co.get("market_cap"),   FMT_DEC1)
        _cell("E", co.get("net_debt_B"),   FMT_DEC1, neg_red=True)
        _cell("F", co.get("ev"),           FMT_DEC1)

        # Financials (in $M)
        rev_m    = co.get("revenue_M")
        ebitda_m = co.get("ebitda_M")
        _cell("G", rev_m,                  FMT_DEC0)
        _cell("H", ebitda_m,               FMT_DEC0, neg_red=True)
        _cell("I", co.get("gross_margin"), FMT_PCT)
        _cell("J", co.get("ebitda_margin"),FMT_PCT)

        # Multiples
        _cell("K", co.get("ev_revenue"),   FMT_MULT, bold=True)
        _cell("L", co.get("ev_ebitda"),    FMT_MULT, bold=True)
        _cell("M", co.get("pe"),           FMT_MULT, bold=True)

    # ── Spacer row ─────────────────────────────────────────────────────────
    spacer_row = data_start + len(companies)
    ws.row_dimensions[spacer_row].height = 6

    # ── Stats rows: High, Low, Median, Mean ────────────────────────────────
    def _safe_list(key):
        return [co[key] for co in companies if co.get(key) is not None]

    stat_defs = [
        ("High",   lambda vals: max(vals) if vals else None),
        ("Low",    lambda vals: min(vals) if vals else None),
        ("Median", lambda vals: median(vals) if vals else None),
        ("Mean",   lambda vals: mean(vals) if vals else None),
    ]

    for j, (label, fn) in enumerate(stat_defs):
        row = spacer_row + 1 + j
        ws.row_dimensions[row].height = 20
        is_last = (j == len(stat_defs) - 1)
        is_med_mean = label in ("Median", "Mean")
        bg = BLUE_LIGHT if is_med_mean else WHITE

        def _stat_cell(col_letter, value, fmt=None, bold=False, h_align="right"):
            c = ws[f"{col_letter}{row}"]
            c.value     = value
            c.font      = _font(bold=bold, size=9, color=DARK_TEXT)
            c.fill      = _fill(bg)
            c.alignment = _align(h_align, "center")
            c.border    = _stat_border(is_last)
            if fmt:
                c.number_format = fmt

        _stat_cell("A", label, bold=True, h_align="left")
        _stat_cell("B", "")
        _stat_cell("C", fn(_safe_list("share_price")), FMT_PRICE)
        _stat_cell("D", fn(_safe_list("market_cap")),  FMT_DEC1)
        _stat_cell("E", fn(_safe_list("net_debt_B")),  FMT_DEC1)
        _stat_cell("F", fn(_safe_list("ev")),          FMT_DEC1)
        _stat_cell("G", fn(_safe_list("revenue_M")),   FMT_DEC0)
        _stat_cell("H", fn(_safe_list("ebitda_M")),    FMT_DEC0)
        _stat_cell("I", fn(_safe_list("gross_margin")),FMT_PCT)
        _stat_cell("J", fn(_safe_list("ebitda_margin")),FMT_PCT)
        _stat_cell("K", fn(_safe_list("ev_revenue")),  FMT_MULT, bold=True)
        _stat_cell("L", fn(_safe_list("ev_ebitda")),   FMT_MULT, bold=True)
        _stat_cell("M", fn(_safe_list("pe")),          FMT_MULT, bold=True)

    # ── Source footer ──────────────────────────────────────────────────────
    src_row = spacer_row + 1 + len(stat_defs) + 1
    ws.row_dimensions[src_row].height = 16
    ws.merge_cells(f"A{src_row}:M{src_row}")
    ws[f"A{src_row}"].value = "Source: Polygon.io, Financial Modeling Prep  |  Generated by tsifl"
    ws[f"A{src_row}"].font  = _font(italic=True, color="9CA3AF", size=8)

    note_row = src_row + 1
    ws.merge_cells(f"A{note_row}:M{note_row}")
    ws[f"A{note_row}"].value = (
        "Note: Market data as of prior close. Financials are LTM (last twelve months). "
        "EV = Market Cap + Net Debt. Multiples calculated using LTM figures."
    )
    ws[f"A{note_row}"].font  = _font(italic=True, color="9CA3AF", size=8)
    ws[f"A{note_row}"].alignment = _align("left", "top", wrap=True)
    ws.row_dimensions[note_row].height = 28

    # ── Freeze panes ───────────────────────────────────────────────────────
    ws.freeze_panes = "C6"

    # ── Print settings ─────────────────────────────────────────────────────
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  POWERPOINT — Single comp slide
# ══════════════════════════════════════════════════════════════════════════════

def generate_comp_slide_pptx(payload: dict) -> bytes:
    """Generate an IB-grade trading comps slide as .pptx bytes."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.oxml import parse_xml

    def rgb(hex_str):
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    companies = payload.get("companies", [])
    title     = payload.get("title", "Trading Comps")
    date_str  = payload.get("date", "")
    currency  = payload.get("currency", "USD")

    # Sort by market cap descending
    companies = sorted(companies, key=lambda c: c.get("market_cap") or 0, reverse=True)

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # White background
    bg = slide.background; bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Title
    box = slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(12.73), Inches(0.55))
    tf = box.text_frame; tf.word_wrap = False
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    run = p.add_run(); run.text = title
    run.font.bold = True; run.font.size = Pt(16)
    run.font.color.rgb = rgb(NAVY); run.font.name = FONT_NAME

    # Subtitle
    box2 = slide.shapes.add_textbox(Inches(0.3), Inches(0.72), Inches(6), Inches(0.3))
    tf2 = box2.text_frame; p2 = tf2.paragraphs[0]
    run2 = p2.add_run()
    run2.text = f"{date_str}  |  {currency}"
    run2.font.size = Pt(8); run2.font.color.rgb = rgb("6B7280")
    run2.font.name = FONT_NAME; run2.font.italic = True

    # Divider
    conn = slide.shapes.add_connector(
        1, Inches(0.3), Inches(0.95), Inches(13.03), Inches(0.95)
    )
    conn.line.color.rgb = rgb(NAVY); conn.line.width = Pt(1.5)

    # ── Table ──────────────────────────────────────────────────────────────
    col_headers = [
        "Company", "Ticker", "Price ($)",
        "Mkt Cap\n($B)", "EV ($B)",
        "Revenue\n($M)", "EBITDA\n($M)", "Gross\nMargin", "EBITDA\nMargin",
        "EV / Rev", "EV /\nEBITDA", "P / E"
    ]
    n_data = len(companies)
    # header + data + spacer + median + mean = n_data + 4
    n_rows = 1 + n_data + 3
    n_cols = len(col_headers)

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols,
        Inches(0.3), Inches(1.05), Inches(12.73), Inches(0.30 * n_rows)
    )
    tbl = tbl_shape.table

    # Column widths (proportional)
    col_widths_in = [2.2, 0.7, 0.8, 0.85, 0.85, 0.9, 0.9, 0.8, 0.8, 0.75, 0.8, 0.7]
    total_w = sum(col_widths_in)
    for ci, w in enumerate(col_widths_in):
        tbl.columns[ci].width = int(Inches(w * 12.73 / total_w))

    def set_cell(ri, ci, text, bold=False, font_size=8,
                 bg_hex=None, fg=DARK_TEXT, align_h=PP_ALIGN.RIGHT, italic=False):
        cell = tbl.cell(ri, ci)
        cell.text = ""
        tf = cell.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]; p.alignment = align_h
        run = p.add_run()
        run.text = str(text) if text is not None else "--"
        run.font.bold = bold; run.font.size = Pt(font_size)
        run.font.color.rgb = rgb(fg); run.font.name = FONT_NAME
        run.font.italic = italic
        if bg_hex:
            fill_xml = f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{bg_hex}"/></a:solidFill>'
            tcPr = cell._tc.get_or_add_tcPr()
            for old in tcPr.findall(qn('a:solidFill')):
                tcPr.remove(old)
            for old in tcPr.findall(qn('a:noFill')):
                tcPr.remove(old)
            tcPr.insert(0, parse_xml(fill_xml))

    def fmt_mult(v):
        if v is None: return "--"
        try: return f"{float(v):.1f}x"
        except: return "--"

    def fmt_dec(v, dp=1):
        if v is None: return "--"
        try: return f"{float(v):,.{dp}f}"
        except: return "--"

    def fmt_pct(v):
        if v is None: return "--"
        try: return f"{float(v)*100:.1f}%"
        except: return "--"

    # Header row
    for ci, h in enumerate(col_headers):
        ha = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER if ci == 1 else PP_ALIGN.RIGHT
        set_cell(0, ci, h, bold=True, font_size=8, bg_hex=NAVY, fg=WHITE, align_h=ha)

    # Data rows
    for i, co in enumerate(companies):
        ri = i + 1
        bg_hex = BLUE_FILL if i % 2 == 1 else WHITE
        set_cell(ri, 0, co.get("name",""),     bg_hex=bg_hex, align_h=PP_ALIGN.LEFT)
        set_cell(ri, 1, co.get("ticker",""),   bg_hex=bg_hex, align_h=PP_ALIGN.CENTER, bold=True)
        set_cell(ri, 2, f"${co['share_price']:.2f}" if co.get("share_price") else "--", bg_hex=bg_hex)
        set_cell(ri, 3, fmt_dec(co.get("market_cap")), bg_hex=bg_hex)
        set_cell(ri, 4, fmt_dec(co.get("ev")),          bg_hex=bg_hex)
        set_cell(ri, 5, fmt_dec(co.get("revenue_M"), 0), bg_hex=bg_hex)
        set_cell(ri, 6, fmt_dec(co.get("ebitda_M"), 0),  bg_hex=bg_hex)
        set_cell(ri, 7, fmt_pct(co.get("gross_margin")), bg_hex=bg_hex)
        set_cell(ri, 8, fmt_pct(co.get("ebitda_margin")),bg_hex=bg_hex)
        set_cell(ri, 9, fmt_mult(co.get("ev_revenue")),  bg_hex=bg_hex, bold=True)
        set_cell(ri,10, fmt_mult(co.get("ev_ebitda")),   bg_hex=bg_hex, bold=True)
        set_cell(ri,11, fmt_mult(co.get("pe")),          bg_hex=bg_hex, bold=True)

    # Spacer
    spacer_ri = 1 + n_data
    for ci in range(n_cols):
        set_cell(spacer_ri, ci, "", bg_hex=WHITE)
    tbl.rows[spacer_ri].height = int(Inches(0.06))

    # Median / Mean
    def stat_vals(key):
        return [co[key] for co in companies if co.get(key) is not None]

    for j, (label, fn) in enumerate([
        ("Median", lambda v: median(v) if v else None),
        ("Mean",   lambda v: mean(v)   if v else None),
    ]):
        ri = spacer_ri + 1 + j
        set_cell(ri, 0, label,     bold=True, bg_hex=BLUE_LIGHT, align_h=PP_ALIGN.LEFT)
        set_cell(ri, 1, "",        bg_hex=BLUE_LIGHT)
        set_cell(ri, 2, "",        bg_hex=BLUE_LIGHT)
        set_cell(ri, 3, fmt_dec(fn(stat_vals("market_cap"))),  bg_hex=BLUE_LIGHT)
        set_cell(ri, 4, fmt_dec(fn(stat_vals("ev"))),          bg_hex=BLUE_LIGHT)
        set_cell(ri, 5, fmt_dec(fn(stat_vals("revenue_M")),0), bg_hex=BLUE_LIGHT)
        set_cell(ri, 6, fmt_dec(fn(stat_vals("ebitda_M")),0),  bg_hex=BLUE_LIGHT)
        set_cell(ri, 7, fmt_pct(fn(stat_vals("gross_margin"))),bg_hex=BLUE_LIGHT)
        set_cell(ri, 8, fmt_pct(fn(stat_vals("ebitda_margin"))),bg_hex=BLUE_LIGHT)
        set_cell(ri, 9, fmt_mult(fn(stat_vals("ev_revenue"))), bg_hex=BLUE_LIGHT, bold=True)
        set_cell(ri,10, fmt_mult(fn(stat_vals("ev_ebitda"))),  bg_hex=BLUE_LIGHT, bold=True)
        set_cell(ri,11, fmt_mult(fn(stat_vals("pe"))),         bg_hex=BLUE_LIGHT, bold=True)

    # Footnote
    fn_box = slide.shapes.add_textbox(Inches(0.3), Inches(7.1), Inches(12.73), Inches(0.3))
    tf3 = fn_box.text_frame; p3 = tf3.paragraphs[0]
    run3 = p3.add_run()
    run3.text = f"Source: Polygon.io, Financial Modeling Prep  |  Generated by tsifl  |  {date_str}"
    run3.font.size = Pt(7); run3.font.color.rgb = rgb("9CA3AF")
    run3.font.name = FONT_NAME; run3.font.italic = True

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  Helper — build payload from live market data
# ══════════════════════════════════════════════════════════════════════════════

async def build_comp_payload(tickers: list[str], title: str = None) -> dict:
    """
    Fetch live data for tickers and build payload for comp generators.

    Data flow:
      1. Polygon → price, market_cap_B, shares
      2. yfinance → revenue_ltm_M, ebitda_ltm_M, margins, net_debt_M
         (FMP fallback if yfinance fails and FMP key is set)
      3. Compute EV, multiples
    """
    import asyncio
    from datetime import datetime
    from services.fmp import get_fundamentals as fmp_get
    from services.yfinance_service import get_fundamentals as yf_get
    from routes.terminal import _build_quote

    upper = [t.upper() for t in tickers]

    # 1. Prices from terminal's cached Polygon layer (shares 5-min cache)
    quotes = await asyncio.gather(*[_build_quote(t) for t in upper])
    mkt = {q["ticker"]: q for q in quotes if q.get("price", 0) > 0}

    # 2. Fundamentals — sequential to avoid Yahoo rate limits on cloud IPs
    #    yfinance has 24h cache built in, so cached tickers are instant
    fund = {}
    for t in upper:
        d = await yf_get(t)
        if d.get("error"):
            d = await fmp_get(t)
        fund[t] = d

    # 3. Build company list
    companies = []
    for ticker in upper:
        m = mkt.get(ticker, {})
        f = fund.get(ticker, {})

        if f.get("error") and not m:
            continue

        price     = _safe_float(m.get("price"))     or _safe_float(f.get("price"))
        mkt_cap_b = _safe_float(m.get("market_cap_B")) or _safe_float(f.get("market_cap_B"))

        rev_m    = _safe_float(f.get("revenue_ltm_M"))
        ebitda_m = _safe_float(f.get("ebitda_ltm_M"))
        ni_m     = _safe_float(f.get("net_income_ltm_M"))
        gm_pct   = _safe_float(f.get("gross_margin_pct"))    # 0–100
        net_debt_m = _safe_float(f.get("net_debt_M"))

        ebitda_margin_pct = (
            round(ebitda_m / rev_m * 100, 1) if ebitda_m and rev_m else None
        )

        # EV = market_cap + net_debt
        net_debt_b = round(net_debt_m / 1e3, 2) if net_debt_m is not None else None
        ev_b = None
        if mkt_cap_b is not None:
            if net_debt_b is not None:
                ev_b = round(mkt_cap_b + net_debt_b, 2)
            else:
                ev_b = mkt_cap_b  # approximate

        # Multiples
        ev_ebitda = round(ev_b * 1e3 / ebitda_m, 1) if ev_b and ebitda_m and ebitda_m > 0 else None
        ev_rev    = round(ev_b * 1e3 / rev_m, 1)    if ev_b and rev_m    and rev_m    > 0 else None
        pe        = round(mkt_cap_b * 1e3 / ni_m, 1) if mkt_cap_b and ni_m and ni_m > 0 else None

        period = f.get("period_label") or "LTM"

        companies.append({
            "ticker":        ticker,
            "name":          f.get("name") or m.get("name") or ticker,
            "period":        period,
            "share_price":   price,
            "market_cap":    mkt_cap_b,                            # $B
            "net_debt_B":    net_debt_b,                           # $B
            "ev":            ev_b,                                 # $B
            "revenue_M":     round(rev_m, 0) if rev_m else None,   # $M
            "ebitda_M":      round(ebitda_m, 0) if ebitda_m else None,
            "gross_margin":  gm_pct / 100 if gm_pct else None,    # decimal
            "ebitda_margin": ebitda_margin_pct / 100 if ebitda_margin_pct else None,
            "ev_ebitda":     ev_ebitda,
            "ev_revenue":    ev_rev,
            "pe":            pe,
        })

    return {
        "title":     title or f"Trading Comps — {', '.join(upper)}",
        "date":      f"As of {datetime.utcnow().strftime('%B %Y')}",
        "currency":  "USD ($M)",
        "companies": companies,
    }


# ══════════════════════════════════════════════════════════════════════════════
#  POWERPOINT — Full 4-slide comp deck
# ══════════════════════════════════════════════════════════════════════════════

def generate_comp_deck_pptx(payload: dict) -> bytes:
    """
    Generate a full 4-slide IB comp deck:
      Slide 1 — Title
      Slide 2 — Trading Comps table
      Slide 3 — Valuation snapshot (KPI boxes)
      Slide 4 — Key takeaways
    """
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import qn
    import copy

    def rgb(hex_str):
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    companies = sorted(
        payload.get("companies", []),
        key=lambda c: c.get("market_cap") or 0, reverse=True
    )
    title    = payload.get("title", "Trading Comps")
    date_str = payload.get("date", "")
    currency = payload.get("currency", "USD")

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def _white_bg(slide):
        bg = slide.background; bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    def _add_text(slide, text, left, top, width, height,
                  bold=False, size=11, color=DARK_TEXT, align=PP_ALIGN.LEFT,
                  italic=False):
        box = slide.shapes.add_textbox(
            Inches(left), Inches(top), Inches(width), Inches(height)
        )
        tf = box.text_frame; tf.word_wrap = True
        p = tf.paragraphs[0]; p.alignment = align
        run = p.add_run(); run.text = text
        run.font.bold = bold; run.font.size = Pt(size)
        run.font.color.rgb = rgb(color); run.font.name = FONT_NAME
        run.font.italic = italic
        return box

    def _divider(slide, y=0.9):
        conn = slide.shapes.add_connector(
            1, Inches(0.3), Inches(y), Inches(13.03), Inches(y)
        )
        conn.line.color.rgb = rgb(NAVY); conn.line.width = Pt(1.2)

    def _header(slide, title_text, sub=""):
        _add_text(slide, title_text, 0.3, 0.18, 12.73, 0.55,
                  bold=True, size=15, color=NAVY)
        if sub:
            _add_text(slide, sub, 0.3, 0.7, 12.73, 0.25,
                      size=8, color="6B7280", italic=True)
        _divider(slide, 0.92)

    def _footnote(slide, text=None):
        note = text or f"Source: Polygon.io, Financial Modeling Prep  |  Generated by tsifl  |  {date_str}"
        _add_text(slide, note, 0.3, 7.1, 12.73, 0.3,
                  size=7, color="9CA3AF", italic=True)

    def fmt_mult(v):
        try: return f"{float(v):.1f}x" if v is not None else "--"
        except: return "--"

    def fmt_pct(v):
        try: return f"{float(v)*100:.1f}%" if v is not None else "--"
        except: return "--"

    def stat_vals(key):
        return [co[key] for co in companies if co.get(key) is not None]

    def med(key):
        vals = stat_vals(key)
        return median(vals) if vals else None

    # ── SLIDE 1 — Title ────────────────────────────────────────────────────
    s1 = prs.slides.add_slide(blank); _white_bg(s1)

    bar = s1.shapes.add_shape(1, Inches(0), Inches(0), Inches(13.33), Inches(2.2))
    bar.fill.solid(); bar.fill.fore_color.rgb = rgb(NAVY)
    bar.line.fill.background()

    _add_text(s1, title, 0.5, 0.45, 12.33, 0.9,
              bold=True, size=32, color=WHITE)
    _add_text(s1, f"Trading Comparables Analysis  |  {date_str}",
              0.5, 1.35, 12.33, 0.4, size=13, color="A8C4E0", italic=True)

    # Ticker chips
    chip_x = 0.5
    for co in companies[:10]:
        chip = s1.shapes.add_shape(1, Inches(chip_x), Inches(2.7), Inches(0.9), Inches(0.35))
        chip.fill.solid(); chip.fill.fore_color.rgb = rgb(NAVY_MED)
        chip.line.color.rgb = rgb("2E5599"); chip.line.width = Pt(0.5)
        tf = chip.text_frame; tf.text = co.get("ticker", "")
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        run = tf.paragraphs[0].runs[0]
        run.font.bold = True; run.font.size = Pt(9)
        run.font.color.rgb = rgb(WHITE); run.font.name = FONT_NAME
        chip_x += 1.05

    _add_text(s1, "Confidential  |  For Discussion Purposes Only",
              0.5, 6.9, 12.33, 0.35, size=8, color="6B7280", italic=True,
              align=PP_ALIGN.CENTER)

    # ── SLIDE 2 — Comp table (clone single slide) ─────────────────────────
    prs2_bytes = generate_comp_slide_pptx(payload)
    prs2 = Presentation(io.BytesIO(prs2_bytes))
    src_slide = prs2.slides[0]
    new_slide = prs.slides.add_slide(blank)
    for shape in src_slide.shapes:
        new_slide.shapes._spTree.insert(2, copy.deepcopy(shape._element))
    _white_bg(new_slide)

    # ── SLIDE 3 — Valuation snapshot ───────────────────────────────────────
    s3 = prs.slides.add_slide(blank); _white_bg(s3)
    _header(s3, "Valuation Snapshot", f"Median multiples  |  {date_str}  |  {currency}")

    kpis = [
        ("EV / EBITDA",  fmt_mult(med("ev_ebitda")),  NAVY),
        ("EV / Revenue", fmt_mult(med("ev_revenue")), NAVY_MED),
        ("P / E",        fmt_mult(med("pe")),         "2E5599"),
    ]
    box_w, box_h = 3.6, 2.4
    spacing = 0.4
    total_w = len(kpis) * box_w + (len(kpis)-1) * spacing
    start_x = (13.33 - total_w) / 2

    for i, (label, value, color) in enumerate(kpis):
        bx = start_x + i * (box_w + spacing)
        bg_shape = s3.shapes.add_shape(1, Inches(bx), Inches(1.8), Inches(box_w), Inches(box_h))
        bg_shape.fill.solid(); bg_shape.fill.fore_color.rgb = rgb(color)
        bg_shape.line.fill.background()
        _add_text(s3, value, bx+0.15, 2.05, box_w-0.3, 1.3,
                  bold=True, size=44, color=WHITE, align=PP_ALIGN.CENTER)
        _add_text(s3, label, bx+0.15, 3.45, box_w-0.3, 0.45,
                  size=11, color="A8C4E0", align=PP_ALIGN.CENTER, italic=True)

    n = len(companies)
    _add_text(s3, f"Based on {n} comparable compan{'y' if n==1 else 'ies'}  |  Median shown",
              0.3, 4.55, 12.73, 0.3, size=8, color="9CA3AF", italic=True, align=PP_ALIGN.CENTER)
    _footnote(s3)

    # ── SLIDE 4 — Key takeaways ────────────────────────────────────────────
    s4 = prs.slides.add_slide(blank); _white_bg(s4)
    _header(s4, "Key Observations", date_str)

    observations = []
    ev_vals = stat_vals("ev_ebitda")
    pe_vals_list = stat_vals("pe")

    if ev_vals:
        hi = max(companies, key=lambda c: c.get("ev_ebitda") or 0)
        lo = min(companies, key=lambda c: c.get("ev_ebitda") or 999)
        observations.append(
            f"EV/EBITDA range: {fmt_mult(lo.get('ev_ebitda'))} ({lo.get('ticker','')}) "
            f"to {fmt_mult(hi.get('ev_ebitda'))} ({hi.get('ticker','')}), "
            f"median {fmt_mult(med('ev_ebitda'))}"
        )
    if pe_vals_list:
        observations.append(
            f"P/E multiples range {fmt_mult(min(pe_vals_list))} to "
            f"{fmt_mult(max(pe_vals_list))}, median {fmt_mult(med('pe'))}"
        )
    margins = stat_vals("ebitda_margin")
    if margins:
        best = max(companies, key=lambda c: c.get("ebitda_margin") or 0)
        observations.append(
            f"{best.get('name', best.get('ticker',''))} leads on EBITDA margin at "
            f"{fmt_pct(best.get('ebitda_margin'))}"
        )
    if len(companies) >= 2:
        by_ev = sorted([c for c in companies if c.get("ev_ebitda")],
                       key=lambda c: c["ev_ebitda"])
        if len(by_ev) >= 2:
            observations.append(
                f"{by_ev[0].get('name', by_ev[0].get('ticker',''))} trades at a discount "
                f"({fmt_mult(by_ev[0].get('ev_ebitda'))} EV/EBITDA) vs. "
                f"{by_ev[-1].get('name', by_ev[-1].get('ticker',''))} "
                f"({fmt_mult(by_ev[-1].get('ev_ebitda'))} EV/EBITDA)"
            )
    if not observations:
        observations = [
            "See trading comps table on prior slide for full detail",
            f"Analysis covers {len(companies)} publicly traded comparable companies",
        ]

    bullet_y = 1.15
    for obs in observations[:5]:
        _add_text(s4, "▸", 0.4, bullet_y, 0.3, 0.45,
                  bold=True, size=10, color=NAVY)
        _add_text(s4, obs, 0.75, bullet_y, 11.8, 0.45,
                  size=10.5, color=DARK_TEXT)
        bullet_y += 0.7

    _footnote(s4)

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  Utility
# ══════════════════════════════════════════════════════════════════════════════

def _safe_float(v, default=None):
    try: return round(float(v), 2) if v is not None else default
    except: return default
