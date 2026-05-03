"""
/generate — server-side file generation endpoints.
Returns perfectly-formatted .xlsx and .pptx files.
No Office.js limitations. No locale bugs.
"""

from fastapi import APIRouter, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional
import io

router = APIRouter()


class CompRequest(BaseModel):
    tickers: Optional[list[str]] = None   # if provided, fetch live data
    payload: Optional[dict]      = None   # if provided, use directly
    title:   Optional[str]       = None


# ── Excel — Trading Comps ─────────────────────────────────────────────────────

@router.post("/comp-table.xlsx")
async def generate_comp_xlsx(req: CompRequest):
    """
    Generate and return a fully-formatted IB comp table as .xlsx.
    Either pass `tickers` (live data fetch) or a full `payload`.
    """
    from services.templates import generate_comp_table_xlsx, build_comp_payload

    if req.tickers:
        payload = await build_comp_payload(req.tickers, title=req.title)
    elif req.payload:
        payload = req.payload
    else:
        raise HTTPException(400, "Provide either tickers or payload")

    if not payload.get("companies"):
        raise HTTPException(
            422,
            "Could not fetch data for any tickers. "
            "Market data APIs may be rate-limited — try again in 60 seconds."
        )

    try:
        xlsx_bytes = generate_comp_table_xlsx(payload)
    except Exception as e:
        raise HTTPException(500, f"Template generation failed: {e}")

    filename = (req.title or "Trading_Comps").replace(" ", "_") + ".xlsx"
    return StreamingResponse(
        io.BytesIO(xlsx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ── PowerPoint — Trading Comps ────────────────────────────────────────────────

@router.post("/comp-slide.pptx")
async def generate_comp_pptx(req: CompRequest):
    """
    Generate and return a fully-formatted IB comp deck (4 slides) as .pptx.
    Either pass `tickers` (live data fetch) or a full `payload`.
    """
    from services.templates import generate_comp_deck_pptx, build_comp_payload

    if req.tickers:
        payload = await build_comp_payload(req.tickers, title=req.title)
    elif req.payload:
        payload = req.payload
    else:
        raise HTTPException(400, "Provide either tickers or payload")

    if not payload.get("companies"):
        raise HTTPException(
            422,
            "Could not fetch data for any tickers. "
            "Market data APIs may be rate-limited — try again in 60 seconds."
        )

    try:
        pptx_bytes = generate_comp_deck_pptx(payload)
    except Exception as e:
        raise HTTPException(500, f"Template generation failed: {e}")

    filename = (req.title or "Trading_Comps").replace(" ", "_") + ".pptx"
    return StreamingResponse(
        io.BytesIO(pptx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ── Inject — returns structured data for direct Excel injection ──────────────

@router.post("/comp-inject")
async def generate_comp_inject(req: CompRequest):
    """
    Build comp data and return it as a 2D array ready to write into Excel.
    If only 1 ticker is given, auto-find peers first.
    Returns: { title, sheet_name, headers, rows, stats, tickers_used }
    """
    from services.templates import build_comp_payload
    from routes.terminal import terminal_peers

    tickers = req.tickers or []
    if not tickers:
        raise HTTPException(400, "Provide at least one ticker")

    # Auto-peer: if user gives 1 ticker, find peers automatically
    if len(tickers) == 1:
        primary = tickers[0].upper()
        peer_data = await terminal_peers(primary)
        peers = peer_data.get("peers", [])[:7]
        tickers = [primary] + peers

    payload = await build_comp_payload(tickers, title=req.title)
    companies = payload.get("companies", [])

    if not companies:
        raise HTTPException(422, "Could not fetch data for any of the provided tickers. Try again in a minute (rate limit may have reset).")

    # Build full rows first, then strip columns that are entirely N/A
    all_headers = [
        "Company", "Ticker", "Price ($)", "Mkt Cap ($B)",
        "Net Debt ($B)", "EV ($B)", "Revenue ($M)", "EBITDA ($M)",
        "Gross %", "EBITDA %", "EV/Rev", "EV/EBITDA", "P/E"
    ]

    def _fmt(v, decimals=1):
        if v is None:
            return "N/A"
        if isinstance(v, (int, float)):
            return round(v, decimals)
        return v

    def _pct(v):
        if v is None:
            return "N/A"
        return round(v * 100, 1)

    full_rows = []
    for c in sorted(companies, key=lambda x: x.get("market_cap") or 0, reverse=True):
        full_rows.append([
            c.get("name", c["ticker"]),
            c["ticker"],
            _fmt(c.get("share_price"), 2),
            _fmt(c.get("market_cap"), 1),
            _fmt(c.get("net_debt_B"), 1),
            _fmt(c.get("ev"), 1),
            _fmt(c.get("revenue_M"), 0),
            _fmt(c.get("ebitda_M"), 0),
            _pct(c.get("gross_margin")),
            _pct(c.get("ebitda_margin")),
            _fmt(c.get("ev_revenue"), 1),
            _fmt(c.get("ev_ebitda"), 1),
            _fmt(c.get("pe"), 1),
        ])

    # Drop columns where EVERY company has N/A (keeps table clean)
    # Always keep: Company (0), Ticker (1) — never dropped
    keep_cols = [0, 1]
    for ci in range(2, len(all_headers)):
        has_data = any(
            isinstance(r[ci], (int, float)) for r in full_rows
        )
        if has_data:
            keep_cols.append(ci)

    headers = [all_headers[i] for i in keep_cols]
    rows = [[r[i] for i in keep_cols] for r in full_rows]

    # Summary stats (High / Low / Median / Mean) for numeric columns
    import statistics
    stat_cols = [i for i, ci in enumerate(keep_cols) if ci >= 2]
    stat_rows = []
    for stat_name in ["High", "Low", "Median", "Mean"]:
        row = ["" for _ in keep_cols]
        row[0] = stat_name
        for si in stat_cols:
            vals = [r[si] for r in rows if isinstance(r[si], (int, float))]
            if not vals:
                row[si] = "N/A"
            elif stat_name == "High":
                row[si] = round(max(vals), 1)
            elif stat_name == "Low":
                row[si] = round(min(vals), 1)
            elif stat_name == "Median":
                row[si] = round(statistics.median(vals), 1)
            elif stat_name == "Mean":
                row[si] = round(statistics.mean(vals), 1)
        stat_rows.append(row)

    primary_ticker = (req.tickers[0] if req.tickers else tickers[0]).upper()

    return {
        "title":        payload.get("title", f"Trading Comps — {primary_ticker}"),
        "sheet_name":   f"Comps {primary_ticker}",
        "date":         payload.get("date", ""),
        "headers":      headers,
        "rows":         rows,
        "stats":        stat_rows,
        "tickers_used": [c["ticker"] for c in companies],
        "count":        len(companies),
    }


# ── Combined — both files in one call (used by Build Deck) ───────────────────

@router.post("/comp-package")
async def generate_comp_package(req: CompRequest):
    """
    Fetch live data for tickers, return URLs for both .xlsx and .pptx.
    The add-in calls this to get pre-formatted downloads ready.
    """
    from services.templates import build_comp_payload

    if not req.tickers:
        raise HTTPException(400, "tickers required")

    payload = await build_comp_payload(req.tickers, title=req.title)
    return {
        "payload":   payload,
        "xlsx_url":  "/generate/comp-table.xlsx",
        "pptx_url":  "/generate/comp-slide.pptx",
        "companies": len(payload["companies"]),
        "title":     payload["title"],
        "date":      payload["date"],
    }
