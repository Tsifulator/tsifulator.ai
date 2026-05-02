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

# ── Shared brand constants ─────────────────────────────────────────────────
NAVY      = "002366"   # IB dark navy — headers, title bars
NAVY_MED  = "1F3864"   # slightly lighter navy — alternating header
BLUE_LIGHT= "D6E4F0"   # light blue — median/mean row bg
BLUE_FILL = "EBF3FB"   # very light — subtle alternating rows
WHITE     = "FFFFFF"
DARK_TEXT = "1A1A2E"
GREY_LINE = "BDC3C7"
GREEN_CHK = "27AE60"
RED_NEG   = "E74C3C"
FONT_NAME = "Calibri"


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL — Trading Comps
# ══════════════════════════════════════════════════════════════════════════════

def generate_comp_table_xlsx(payload: dict) -> bytes:
    """
    Generate a fully-formatted IB trading comps table as .xlsx bytes.

    payload = {
      "title": str,
      "date":  str,          # "As of May 2026"
      "currency": str,       # "USD ($M)"
      "companies": [
        { "ticker", "name", "period",
          "revenue", "gross_margin", "ebitda", "ebitda_margin",
          "share_price", "market_cap", "ev",
          "pe", "ev_ebitda", "ev_revenue" }
      ]
    }
    """
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
    from statistics import median, mean

    wb = Workbook()
    ws = wb.active
    ws.title = "Trading Comps"

    companies = payload.get("companies", [])
    title     = payload.get("title", "Trading Comps")
    date_str  = payload.get("date", "")
    currency  = payload.get("currency", "USD ($M)")

    # ── Style helpers ─────────────────────────────────────────────────────
    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def font(bold=False, color=DARK_TEXT, size=10, name=FONT_NAME, italic=False):
        return Font(bold=bold, color=color, size=size, name=name, italic=italic)

    def align(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    thin  = Side(style="thin",   color=GREY_LINE)
    thick = Side(style="medium", color=NAVY)

    def border(top=None, bottom=None, left=None, right=None):
        return Border(top=top, bottom=bottom, left=left, right=right)

    def all_border(style="thin"):
        s = Side(style=style, color=GREY_LINE)
        return Border(top=s, bottom=s, left=s, right=s)

    def header_border():
        return Border(
            top=Side(style="medium", color=NAVY),
            bottom=Side(style="medium", color=NAVY),
            left=Side(style="thin", color=WHITE),
            right=Side(style="thin", color=WHITE),
        )

    # ── Column layout ──────────────────────────────────────────────────────
    # A=Ticker, B=Company, C=Period, D=Revenue, E=Gross%, F=EBITDA,
    # G=EBITDA%, H=EV, I=Mkt Cap, J=EV/EBITDA, K=EV/Rev, L=P/E
    cols = ["A","B","C","D","E","F","G","H","I","J","K","L"]
    col_widths = [10, 26, 10, 12, 12, 12, 12, 12, 12, 12, 10, 10]
    for i, (col, w) in enumerate(zip(cols, col_widths)):
        ws.column_dimensions[col].width = w

    # Row heights
    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[4].height = 32

    # ── Row 1 — Title ──────────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    c = ws["A1"]
    c.value         = title
    c.font          = Font(bold=True, color=WHITE, size=14, name=FONT_NAME)
    c.fill          = fill(NAVY)
    c.alignment     = align("left", "center")
    c.border        = border(bottom=thick)

    # ── Row 2 — Subtitle ──────────────────────────────────────────────────
    ws.merge_cells("A2:C2")
    ws["A2"].value     = date_str
    ws["A2"].font      = font(italic=True, color="6B7280", size=9)
    ws["A2"].alignment = align("left", "center")

    ws.merge_cells("D2:L2")
    ws["D2"].value     = currency
    ws["D2"].font      = font(italic=True, color="6B7280", size=9)
    ws["D2"].alignment = align("right", "center")

    # ── Row 3 — blank ─────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 6

    # ── Row 4 — Column headers ────────────────────────────────────────────
    headers = [
        "Ticker", "Company", "Period",
        "Revenue", "Gross Margin", "EBITDA", "EBITDA Margin",
        "EV ($B)", "Mkt Cap ($B)",
        "EV / EBITDA", "EV / Revenue", "P / E"
    ]
    for col_letter, header_text in zip(cols, headers):
        c = ws[f"{col_letter}4"]
        c.value     = header_text
        c.font      = Font(bold=True, color=WHITE, size=9, name=FONT_NAME)
        c.fill      = fill(NAVY)
        c.alignment = align("center", "center", wrap=True)
        c.border    = header_border()

    # ── Data rows ──────────────────────────────────────────────────────────
    FMT_INT    = '#,##0'
    FMT_DEC1   = '#,##0.0'
    FMT_MULT   = '#,##0.0x'
    FMT_PCT    = '0.0%'
    FMT_PRICE  = '[$$-409]#,##0.00'

    ev_ebitda_vals, ev_rev_vals, pe_vals = [], [], []
    ebitda_margin_vals, gross_margin_vals = [], []

    for i, co in enumerate(companies):
        row = 5 + i
        ws.row_dimensions[row].height = 18
        row_fill = fill(BLUE_FILL) if i % 2 == 1 else fill(WHITE)

        def cell(col_letter, value, fmt=None, h_align="right"):
            c = ws[f"{col_letter}{row}"]
            c.value     = value
            c.font      = font(size=9)
            c.fill      = row_fill
            c.alignment = align(h_align, "center")
            c.border    = all_border()
            if fmt:
                c.number_format = fmt

        # Text cols
        cell("A", co.get("ticker",""), h_align="center")
        cell("B", co.get("name",""),   h_align="left")
        cell("C", co.get("period",""), h_align="center")
        ws[f"A{row}"].font = Font(bold=True, size=9, name=FONT_NAME)

        # Numeric cols
        cell("D", co.get("revenue"),       FMT_DEC1)
        cell("E", co.get("gross_margin"),  FMT_PCT)
        cell("F", co.get("ebitda"),        FMT_DEC1)
        cell("G", co.get("ebitda_margin"), FMT_PCT)
        cell("H", co.get("ev"),            FMT_DEC1)
        cell("I", co.get("market_cap"),    FMT_DEC1)
        cell("J", co.get("ev_ebitda"),     FMT_MULT)
        cell("K", co.get("ev_revenue"),    FMT_MULT)
        cell("L", co.get("pe"),            FMT_MULT)

        # Collect for stats
        if co.get("ev_ebitda"):  ev_ebitda_vals.append(co["ev_ebitda"])
        if co.get("ev_revenue"): ev_rev_vals.append(co["ev_revenue"])
        if co.get("pe"):         pe_vals.append(co["pe"])
        if co.get("ebitda_margin"): ebitda_margin_vals.append(co["ebitda_margin"])
        if co.get("gross_margin"):  gross_margin_vals.append(co["gross_margin"])

    # ── Median / Mean rows ─────────────────────────────────────────────────
    stat_rows = [
        ("Median", lambda vals: median(vals) if vals else None),
        ("Mean",   lambda vals: mean(vals)   if vals else None),
    ]
    blank_row  = 5 + len(companies)
    ws.row_dimensions[blank_row].height = 8

    for j, (label, fn) in enumerate(stat_rows):
        row = 5 + len(companies) + 1 + j
        ws.row_dimensions[row].height = 18

        def stat_cell(col_letter, value, fmt=None, bold=False, h="right"):
            c = ws[f"{col_letter}{row}"]
            c.value     = value
            c.font      = Font(bold=bold, color=DARK_TEXT, size=9, name=FONT_NAME)
            c.fill      = fill(BLUE_LIGHT)
            c.alignment = align(h, "center")
            c.border    = Border(
                top=Side(style="thin", color=NAVY),
                bottom=Side(style="thin", color=NAVY),
                left=Side(style="thin", color=GREY_LINE),
                right=Side(style="thin", color=GREY_LINE),
            )
            if fmt: c.number_format = fmt

        stat_cell("A", label,                bold=True, h="center")
        stat_cell("B", "",                   h="left")
        stat_cell("C", "")
        stat_cell("D", fn([c.get("revenue") for c in companies if c.get("revenue")]),      FMT_DEC1)
        stat_cell("E", fn(gross_margin_vals),  FMT_PCT)
        stat_cell("F", fn([c.get("ebitda") for c in companies if c.get("ebitda")]),        FMT_DEC1)
        stat_cell("G", fn(ebitda_margin_vals), FMT_PCT)
        stat_cell("H", fn([c.get("ev") for c in companies if c.get("ev")]),                FMT_DEC1)
        stat_cell("I", fn([c.get("market_cap") for c in companies if c.get("market_cap")]),FMT_DEC1)
        stat_cell("J", fn(ev_ebitda_vals),    FMT_MULT, bold=True)
        stat_cell("K", fn(ev_rev_vals),       FMT_MULT, bold=True)
        stat_cell("L", fn(pe_vals),           FMT_MULT, bold=True)

    # ── Sources & Methodology ──────────────────────────────────────────────
    src_row = 5 + len(companies) + 4
    ws.row_dimensions[src_row].height = 14
    ws.merge_cells(f"A{src_row}:L{src_row}")
    ws[f"A{src_row}"].value     = "Sources & Methodology"
    ws[f"A{src_row}"].font      = Font(bold=True, color=DARK_TEXT, size=8, name=FONT_NAME)
    ws[f"A{src_row}"].alignment = align("left", "center")

    for k, co in enumerate(companies):
        r = src_row + 1 + k
        ws.row_dimensions[r].height = 13
        ws.merge_cells(f"A{r}:C{r}")
        ws[f"A{r}"].value     = co.get("ticker","")
        ws[f"A{r}"].font      = Font(bold=True, size=8, name=FONT_NAME)
        ws[f"A{r}"].alignment = align("left", "center")
        ws.merge_cells(f"D{r}:H{r}")
        ws[f"D{r}"].value     = "Financial Modeling Prep (FMP)"
        ws[f"D{r}"].font      = Font(size=8, name=FONT_NAME, color="6B7280")
        ws[f"I{r}"].value     = co.get("period","")
        ws[f"I{r}"].font      = Font(size=8, name=FONT_NAME, color="6B7280")

    # ── Freeze panes at row 5 ─────────────────────────────────────────────
    ws.freeze_panes = "D5"

    # ── Print settings ─────────────────────────────────────────────────────
    ws.page_setup.orientation  = "landscape"
    ws.page_setup.fitToPage    = True
    ws.page_setup.fitToWidth   = 1
    ws.page_setup.fitToHeight  = 0

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  POWERPOINT — Trading Comps Slide
# ══════════════════════════════════════════════════════════════════════════════

def generate_comp_slide_pptx(payload: dict) -> bytes:
    """
    Generate an IB-grade trading comps PPT slide as .pptx bytes.
    Returns a full presentation (1 comp slide) — append more slides as needed.
    """
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.oxml.ns import qn
    from pptx.util import Inches, Pt
    import copy
    from lxml import etree
    from statistics import median, mean

    def rgb(hex_str):
        h = hex_str.lstrip("#")
        return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))

    companies = payload.get("companies", [])
    title     = payload.get("title", "Trading Comps")
    date_str  = payload.get("date", "")
    currency  = payload.get("currency", "USD")

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(blank_layout)

    # ── Title bar ─────────────────────────────────────────────────────────
    from pptx.util import Inches, Pt, Emu
    title_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(0.2), Inches(12.73), Inches(0.55)
    )
    tf = title_box.text_frame
    tf.word_wrap = False
    p  = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    run = p.add_run()
    run.text = title
    run.font.bold   = True
    run.font.size   = Pt(16)
    run.font.color.rgb = rgb(NAVY)
    run.font.name   = "Calibri"

    # Date subtitle
    date_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(0.72), Inches(6), Inches(0.3)
    )
    tf2 = date_box.text_frame
    p2  = tf2.paragraphs[0]
    run2 = p2.add_run()
    run2.text = f"{date_str}  |  {currency}"
    run2.font.size  = Pt(8)
    run2.font.color.rgb = rgb("6B7280")
    run2.font.name  = "Calibri"
    run2.font.italic = True

    # Thin navy divider line under title
    from pptx.util import Emu
    from pptx.oxml import parse_xml
    connector = slide.shapes.add_connector(
        1,  # MSO_CONNECTOR_TYPE.STRAIGHT
        Inches(0.3), Inches(0.95),
        Inches(13.03), Inches(0.95)
    )
    connector.line.color.rgb = rgb(NAVY)
    connector.line.width = Pt(1.5)

    # ── Build table ────────────────────────────────────────────────────────
    col_headers = [
        "Ticker", "Company", "Period",
        "Revenue", "Gross\nMargin", "EBITDA", "EBITDA\nMargin",
        "EV ($B)", "Mkt Cap\n($B)", "EV /\nEBITDA", "EV / Rev", "P / E"
    ]
    col_widths_in = [0.75, 2.2, 0.75, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.85, 0.75, 0.7]

    n_data   = len(companies)
    n_rows   = 1 + n_data + 2 + 1  # header + data + median/mean + spacer
    n_cols   = len(col_headers)

    tbl_top    = Inches(1.05)
    tbl_left   = Inches(0.3)
    tbl_width  = Inches(12.73)
    tbl_height = Inches(0.32 * n_rows)

    tbl_shape = slide.shapes.add_table(
        n_rows, n_cols, tbl_left, tbl_top, tbl_width, tbl_height
    )
    tbl = tbl_shape.table

    # Column widths
    total_w_in = sum(col_widths_in)
    for ci, w in enumerate(col_widths_in):
        tbl.columns[ci].width = int(Inches(w * 12.73 / total_w_in))

    def set_cell(ri, ci, text, bold=False, font_size=8,
                 bg=None, fg=DARK_TEXT, align=PP_ALIGN.RIGHT,
                 italic=False):
        cell = tbl.cell(ri, ci)
        cell.text = ""
        tf = cell.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = str(text) if text is not None else "—"
        run.font.bold   = bold
        run.font.size   = Pt(font_size)
        run.font.color.rgb = rgb(fg)
        run.font.name   = "Calibri"
        run.font.italic = italic
        if bg:
            fill_xml = f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{bg}"/></a:solidFill>'
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            # remove existing fill
            for old in tcPr.findall(qn('a:solidFill')):
                tcPr.remove(old)
            for old in tcPr.findall(qn('a:noFill')):
                tcPr.remove(old)
            tcPr.insert(0, parse_xml(fill_xml))

    def fmt_mult(v):
        if v is None: return "—"
        try: return f"{float(v):.1f}x"
        except: return "—"

    def fmt_dec(v, dp=1):
        if v is None: return "—"
        try: return f"{float(v):,.{dp}f}"
        except: return "—"

    def fmt_pct(v):
        if v is None: return "—"
        try: return f"{float(v)*100:.1f}%"
        except: return "—"

    # Header row
    for ci, h in enumerate(col_headers):
        halign = PP_ALIGN.CENTER if ci < 3 else PP_ALIGN.RIGHT
        set_cell(0, ci, h, bold=True, font_size=8,
                 bg=NAVY, fg=WHITE, align=halign)

    # Data rows
    for i, co in enumerate(companies):
        ri  = i + 1
        bg  = BLUE_FILL if i % 2 == 1 else WHITE
        set_cell(ri, 0, co.get("ticker",""),          bold=True, bg=bg, align=PP_ALIGN.CENTER)
        set_cell(ri, 1, co.get("name",""),            bg=bg, align=PP_ALIGN.LEFT)
        set_cell(ri, 2, co.get("period",""),          bg=bg, align=PP_ALIGN.CENTER, italic=True)
        set_cell(ri, 3, fmt_dec(co.get("revenue")),   bg=bg)
        set_cell(ri, 4, fmt_pct(co.get("gross_margin")), bg=bg)
        set_cell(ri, 5, fmt_dec(co.get("ebitda")),    bg=bg)
        set_cell(ri, 6, fmt_pct(co.get("ebitda_margin")), bg=bg)
        set_cell(ri, 7, fmt_dec(co.get("ev")),        bg=bg)
        set_cell(ri, 8, fmt_dec(co.get("market_cap")),bg=bg)
        set_cell(ri, 9, fmt_mult(co.get("ev_ebitda")),bg=bg, bold=True)
        set_cell(ri,10, fmt_mult(co.get("ev_revenue")),bg=bg)
        set_cell(ri,11, fmt_mult(co.get("pe")),       bg=bg)

    # Spacer row
    spacer_ri = 1 + n_data
    for ci in range(n_cols):
        set_cell(spacer_ri, ci, "", bg=WHITE)
    tbl.rows[spacer_ri].height = int(Inches(0.08))

    # Median / Mean
    def stat_vals(key):
        return [co[key] for co in companies if co.get(key) is not None]

    stat_defs = [
        ("Median", lambda k: median(stat_vals(k)) if stat_vals(k) else None),
        ("Mean",   lambda k: mean(stat_vals(k))   if stat_vals(k) else None),
    ]
    for j, (label, fn) in enumerate(stat_defs):
        ri = 1 + n_data + 1 + j
        set_cell(ri, 0, label,        bold=True, bg=BLUE_LIGHT, align=PP_ALIGN.CENTER)
        set_cell(ri, 1, "",           bg=BLUE_LIGHT, align=PP_ALIGN.LEFT)
        set_cell(ri, 2, "",           bg=BLUE_LIGHT)
        set_cell(ri, 3, fmt_dec(fn("revenue")),       bg=BLUE_LIGHT)
        set_cell(ri, 4, fmt_pct(fn("gross_margin")),  bg=BLUE_LIGHT)
        set_cell(ri, 5, fmt_dec(fn("ebitda")),        bg=BLUE_LIGHT)
        set_cell(ri, 6, fmt_pct(fn("ebitda_margin")), bg=BLUE_LIGHT)
        set_cell(ri, 7, fmt_dec(fn("ev")),            bg=BLUE_LIGHT)
        set_cell(ri, 8, fmt_dec(fn("market_cap")),    bg=BLUE_LIGHT)
        set_cell(ri, 9, fmt_mult(fn("ev_ebitda")),    bg=BLUE_LIGHT, bold=True)
        set_cell(ri,10, fmt_mult(fn("ev_revenue")),   bg=BLUE_LIGHT)
        set_cell(ri,11, fmt_mult(fn("pe")),           bg=BLUE_LIGHT)

    # ── Footnote ──────────────────────────────────────────────────────────
    fn_box = slide.shapes.add_textbox(
        Inches(0.3), Inches(7.1), Inches(12.73), Inches(0.3)
    )
    tf3 = fn_box.text_frame
    p3  = tf3.paragraphs[0]
    run3 = p3.add_run()
    run3.text = f"Source: Financial Modeling Prep (FMP), Polygon.io  |  Generated by tsifl  |  {date_str}"
    run3.font.size   = Pt(7)
    run3.font.color.rgb = rgb("9CA3AF")
    run3.font.name   = "Calibri"
    run3.font.italic = True

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  Helper — build payload from live market data
# ══════════════════════════════════════════════════════════════════════════════

async def build_comp_payload(tickers: list[str], title: str = None) -> dict:
    """
    Fetch live data for a list of tickers and return a payload
    ready for generate_comp_table_xlsx / generate_comp_slide_pptx.
    """
    from datetime import datetime
    from services.fmp import get_fundamentals as fmp_get
    from services.yfinance_service import get_fundamentals as yf_get

    companies = []
    for ticker in tickers:
        data = await fmp_get(ticker.upper())
        if data.get("error"):
            data = await yf_get(ticker.upper())
        if data.get("error"):
            continue

        companies.append({
            "ticker":        ticker.upper(),
            "name":          data.get("name", ticker),
            "period":        data.get("period", "LTM"),
            "revenue":       _safe_float(data.get("revenue")),
            "gross_margin":  _safe_pct(data.get("gross_margin")),
            "ebitda":        _safe_float(data.get("ebitda")),
            "ebitda_margin": _safe_pct(data.get("ebitda_margin")),
            "share_price":   _safe_float(data.get("price")),
            "market_cap":    _safe_billions(data.get("market_cap")),
            "ev":            _safe_billions(data.get("ev")),
            "pe":            _safe_float(data.get("pe")),
            "ev_ebitda":     _safe_float(data.get("ev_ebitda")),
            "ev_revenue":    _safe_float(data.get("ev_revenue")),
        })

    return {
        "title":    title or f"Trading Comps — {', '.join(tickers)}",
        "date":     f"As of {datetime.utcnow().strftime('%B %Y')}",
        "currency": "USD ($B)",
        "companies": companies,
    }


def _safe_float(v, default=None):
    try: return round(float(v), 2) if v is not None else default
    except: return default

def _safe_pct(v, default=None):
    """Return as decimal (0.35 not 35) — formatters handle display."""
    try:
        f = float(v)
        return f / 100 if f > 1 else f  # normalise if stored as 35.0
    except: return default

def _safe_billions(v, default=None):
    try:
        f = float(v)
        return round(f / 1e9, 2) if f > 1e6 else round(f, 2)
    except: return default
