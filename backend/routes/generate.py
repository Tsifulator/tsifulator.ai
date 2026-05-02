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
