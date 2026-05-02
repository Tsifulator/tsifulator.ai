"""
/terminal — market data endpoints for the tsifl terminal.
Uses Polygon for prices/history (reliable on cloud), yfinance for fundamentals only.

Server-side cache: Polygon prev-close data only updates once/day, so we cache
aggressively (5 min TTL) to stay under the free-tier 5 calls/min limit.
"""

from fastapi import APIRouter, HTTPException
import os, logging, time, asyncio, hashlib, json
from datetime import datetime, timedelta

router = APIRouter()
logger = logging.getLogger(__name__)

POLYGON_BASE = "https://api.polygon.io"

# ── In-memory cache ─────────────────────────────────────────────────────────
_cache: dict[str, tuple[float, dict]] = {}   # key → (expires_at, data)
CACHE_TTL = 300   # 5 minutes — prev-close only changes once/day


def _cache_key(path: str, params: dict) -> str:
    """Deterministic cache key from path + sorted params (minus apiKey)."""
    clean = {k: v for k, v in sorted(params.items()) if k != "apiKey"}
    raw = f"{path}|{json.dumps(clean, sort_keys=True)}"
    return hashlib.md5(raw.encode()).hexdigest()


def _cache_get(key: str) -> dict | None:
    entry = _cache.get(key)
    if entry and entry[0] > time.time():
        return entry[1]
    return None


def _cache_set(key: str, data: dict, ttl: int = CACHE_TTL):
    _cache[key] = (time.time() + ttl, data)
    # Evict old entries if cache grows too large (>500 entries)
    if len(_cache) > 500:
        now = time.time()
        expired = [k for k, (exp, _) in _cache.items() if exp <= now]
        for k in expired:
            del _cache[k]


# ── Polygon helpers ──────────────────────────────────────────────────────────

async def _polygon_get(path: str, params: dict = {}, ttl: int = CACHE_TTL) -> dict:
    """Generic Polygon API GET with API key injection + cache."""
    key = _cache_key(path, params)
    cached = _cache_get(key)
    if cached is not None:
        return cached

    api_key = os.getenv("POLYGON_API_KEY", "")
    if not api_key:
        return {"error": "POLYGON_API_KEY not set"}
    import httpx
    params = {**params, "apiKey": api_key}
    try:
        async with httpx.AsyncClient(timeout=15.0) as client:
            r = await client.get(f"{POLYGON_BASE}{path}", params=params)
            data = r.json()
            # Only cache successful responses
            if r.status_code == 200 and "error" not in data:
                _cache_set(key, data, ttl)
            return data
    except Exception as e:
        logger.warning(f"[polygon] {path} failed: {e}")
        return {"error": str(e)}


# ── Quote (single) ──────────────────────────────────────────────────────────

async def _build_quote(ticker: str) -> dict:
    """Build quote dict for a single ticker (used by both single + batch)."""
    prev_task = _polygon_get(f"/v2/aggs/ticker/{ticker}/prev", {"adjusted": "true"})
    ref_task  = _polygon_get(f"/v3/reference/tickers/{ticker}")
    prev_data, ref_data = await asyncio.gather(prev_task, ref_task)

    # Previous close bar
    results = prev_data.get("results", [])
    bar = results[0] if results else {}
    price      = bar.get("c", 0)
    open_price = bar.get("o", 0)
    volume     = bar.get("v")

    # Reference data
    ref = ref_data.get("results", {})
    shares = (
        ref.get("weighted_shares_outstanding")
        or ref.get("share_class_shares_outstanding")
        or 0
    )
    market_cap_b = round((price * shares) / 1e9, 2) if shares and price else None
    name = ref.get("name", ticker)
    sic  = ref.get("sic_description", "")

    return {
        "ticker":       ticker,
        "name":         name,
        "price":        round(price, 2),
        "change":       round(price - open_price, 2) if open_price else 0,
        "change_pct":   round(((price - open_price) / open_price) * 100, 2) if open_price and open_price != 0 else 0,
        "market_cap_B": market_cap_b,
        "volume":       int(volume) if volume is not None and volume > 0 else None,
        "sector":       sic or None,
        "industry":     sic or None,
        "currency":     "USD",
    }


@router.get("/quote/{ticker}")
async def terminal_quote(ticker: str):
    """Single-ticker quote: price, change, market cap, volume."""
    return await _build_quote(ticker.upper().strip())


# ── Batch quotes (all watchlist tickers in one HTTP call) ────────────────────

@router.get("/batch-quotes")
async def terminal_batch_quotes(tickers: str = ""):
    """
    Fetch quotes for multiple tickers in one call.
    Usage: /terminal/batch-quotes?tickers=AAPL,MSFT,GOOGL
    Cached data returns instantly; only uncached tickers hit Polygon.
    """
    ticker_list = [t.strip().upper() for t in tickers.split(",") if t.strip()]
    if not ticker_list:
        return {"results": {}}

    # Cap at 15 to prevent abuse
    ticker_list = ticker_list[:15]

    # Run all quote builds concurrently — cache makes this nearly free
    tasks = [_build_quote(t) for t in ticker_list]
    quotes = await asyncio.gather(*tasks, return_exceptions=True)

    results = {}
    for t, q in zip(ticker_list, quotes):
        if isinstance(q, Exception):
            results[t] = {"ticker": t, "error": str(q)}
        else:
            results[t] = q

    return {"results": results}


# ── Fundamentals ─────────────────────────────────────────────────────────────

@router.get("/fundamentals/{ticker}")
async def terminal_fundamentals(ticker: str):
    """Fundamental metrics. Uses yfinance financial statements (not info dict)."""
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
    OHLCV history via Polygon aggregates.
    period: 5d | 1mo | 6mo | 1y | 5y
    interval: 1d | 1wk
    """
    ticker = ticker.upper().strip()

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
            {"adjusted": "true", "sort": "asc", "limit": "5000"},
            ttl=300,  # 5-min cache for history too
        )

        raw_bars = data.get("results", [])
        bars = []
        for b in raw_bars:
            ts = b.get("t", 0) // 1000  # ms → seconds for TradingView
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
    """Ticker/company search via Polygon reference."""
    try:
        data = await _polygon_get("/v3/reference/tickers", {
            "search": query,
            "active": "true",
            "market": "stocks",
            "limit":  "8",
        }, ttl=600)  # cache search results 10 min
        results = data.get("results", [])
        return {"results": [
            {
                "ticker":   r.get("ticker", ""),
                "name":     r.get("name", ""),
                "exchange": r.get("primary_exchange", ""),
                "type":     r.get("type", ""),
            }
            for r in results
        ]}
    except Exception:
        return {"results": []}
