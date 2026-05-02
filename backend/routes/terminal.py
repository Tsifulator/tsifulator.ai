"""
/terminal — market data endpoints for the tsifl terminal.
Uses Polygon for prices/history (reliable on cloud), yfinance for fundamentals only.
"""

from fastapi import APIRouter, HTTPException
import os, logging
from datetime import datetime, timedelta

router = APIRouter()
logger = logging.getLogger(__name__)

POLYGON_BASE = "https://api.polygon.io"


# ── Polygon helpers ──────────────────────────────────────────────────────────

async def _polygon_get(path: str, params: dict = {}) -> dict:
    """Generic Polygon API GET with API key injection."""
    api_key = os.getenv("POLYGON_API_KEY", "")
    if not api_key:
        return {"error": "POLYGON_API_KEY not set"}
    import httpx
    params["apiKey"] = api_key
    async with httpx.AsyncClient(timeout=15.0) as client:
        r = await client.get(f"{POLYGON_BASE}{path}", params=params)
        return r.json()


# ── Quote ────────────────────────────────────────────────────────────────────

@router.get("/quote/{ticker}")
async def terminal_quote(ticker: str):
    """
    Single-ticker quote: price, change, market cap, volume.
    All from Polygon — no yfinance (yfinance.info is blocked on Railway).
    """
    ticker = ticker.upper().strip()
    import asyncio

    # Fetch previous close + reference in parallel
    prev_task = _polygon_get(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
    ref_task  = _polygon_get(f"/v3/reference/tickers/{ticker}")

    prev_data, ref_data = await asyncio.gather(prev_task, ref_task)

    # Previous close
    results = prev_data.get("results", [])
    bar = results[0] if results else {}
    price      = bar.get("c", 0)
    open_price = bar.get("o", 0)
    volume     = bar.get("v", 0)
    prev_close = bar.get("c", price)  # prev day close

    # Reference data
    ref = ref_data.get("results", {})
    shares = (
        ref.get("weighted_shares_outstanding")
        or ref.get("share_class_shares_outstanding")
        or 0
    )
    market_cap_b = round((price * shares) / 1e9, 2) if shares else None
    name = ref.get("name", ticker)
    sic  = ref.get("sic_description", "")

    # 52-week high/low from Polygon aggregate (1 year of daily bars, just min/max)
    week52_high = None
    week52_low  = None
    try:
        today = datetime.utcnow()
        yr_ago = today - timedelta(days=365)
        agg = await _polygon_get(
            f"/v2/aggs/ticker/{ticker}/range/1/day/{yr_ago.strftime('%Y-%m-%d')}/{today.strftime('%Y-%m-%d')}",
            {"adjusted": "true", "sort": "asc", "limit": "365"}
        )
        agg_bars = agg.get("results", [])
        if agg_bars:
            highs = [b["h"] for b in agg_bars if "h" in b]
            lows  = [b["l"] for b in agg_bars if "l" in b]
            week52_high = round(max(highs), 2) if highs else None
            week52_low  = round(min(lows), 2)  if lows  else None
            # Get avg volume from bars
            vols = [b.get("v", 0) for b in agg_bars]
            avg_volume = int(sum(vols) / len(vols)) if vols else None
        else:
            avg_volume = None
    except Exception:
        avg_volume = None

    return {
        "ticker":       ticker,
        "name":         name,
        "price":        round(price, 2),
        "change":       round(price - open_price, 2) if open_price else 0,
        "change_pct":   round(((price - open_price) / open_price) * 100, 2) if open_price and open_price != 0 else 0,
        "market_cap_B": market_cap_b,
        "volume":       int(volume) if volume else None,
        "avg_volume":   avg_volume,
        "pe":           None,   # needs fundamentals — filled by frontend from /fundamentals
        "eps":          None,
        "week52_high":  week52_high,
        "week52_low":   week52_low,
        "sector":       sic or None,
        "industry":     sic or None,
        "currency":     "USD",
    }


# ── Fundamentals ─────────────────────────────────────────────────────────────

@router.get("/fundamentals/{ticker}")
async def terminal_fundamentals(ticker: str):
    """
    Fundamental metrics. Uses yfinance financial statements (not info dict).
    """
    ticker = ticker.upper().strip()
    try:
        from services.yfinance_service import get_fundamentals
        data = await get_fundamentals(ticker)
        if data.get("error"):
            return {"ticker": ticker, "error": data["error"]}
        return data
    except Exception as e:
        return {"ticker": ticker, "error": str(e)}


# ── History ──────────────────────────────────────────────────────────────────

@router.get("/history/{ticker}")
async def terminal_history(ticker: str, period: str = "1y", interval: str = "1d"):
    """
    OHLCV history via Polygon aggregates (reliable on cloud, unlike yfinance).
    period: 5d | 1mo | 6mo | 1y | 5y
    interval: 1d | 1wk
    """
    ticker = ticker.upper().strip()
    api_key = os.getenv("POLYGON_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "POLYGON_API_KEY not set")

    # Map period to date range
    today = datetime.utcnow()
    period_map = {
        "5d":  timedelta(days=7),
        "1mo": timedelta(days=35),
        "3mo": timedelta(days=95),
        "6mo": timedelta(days=185),
        "1y":  timedelta(days=370),
        "2y":  timedelta(days=740),
        "5y":  timedelta(days=1850),
    }
    delta = period_map.get(period, timedelta(days=370))
    from_date = (today - delta).strftime("%Y-%m-%d")
    to_date   = today.strftime("%Y-%m-%d")

    # Map interval to Polygon timespan
    timespan = "day"
    multiplier = 1
    if interval == "1wk":
        timespan = "week"
    elif interval in ("15m", "5m", "1m"):
        timespan = "minute"
        multiplier = int(interval.replace("m", ""))
    elif interval == "1h":
        timespan = "hour"

    try:
        data = await _polygon_get(
            f"/v2/aggs/ticker/{ticker}/range/{multiplier}/{timespan}/{from_date}/{to_date}",
            {"adjusted": "true", "sort": "asc", "limit": "5000"}
        )

        raw_bars = data.get("results", [])
        bars = []
        for b in raw_bars:
            ts = b.get("t", 0) // 1000  # Polygon returns ms, TradingView wants seconds
            bars.append({
                "time":   ts,
                "open":   round(b.get("o", 0), 4),
                "high":   round(b.get("h", 0), 4),
                "low":    round(b.get("l", 0), 4),
                "close":  round(b.get("c", 0), 4),
                "volume": int(b.get("v", 0)),
            })

        return {"ticker": ticker, "period": period, "interval": interval, "bars": bars}

    except Exception as e:
        logger.warning(f"[terminal] history {ticker} failed: {e}")
        raise HTTPException(500, str(e))


# ── Search ───────────────────────────────────────────────────────────────────

@router.get("/search/{query}")
async def terminal_search(query: str):
    """
    Ticker/company search via Polygon reference.
    """
    try:
        data = await _polygon_get("/v3/reference/tickers", {
            "search": query,
            "active": "true",
            "market": "stocks",
            "limit":  "8",
        })
        results = data.get("results", [])
        out = []
        for r in results:
            out.append({
                "ticker":   r.get("ticker", ""),
                "name":     r.get("name", ""),
                "exchange": r.get("primary_exchange", ""),
                "type":     r.get("type", ""),
            })
        return {"results": out}
    except Exception:
        return {"results": []}
