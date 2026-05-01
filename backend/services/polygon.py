"""
Polygon.io service — real-time and previous-close stock data.
Eliminates manual price lookup for comp builders.

Free tier:  previous close data, 5 calls/min
Starter:    $29/mo — real-time snapshots, unlimited calls
"""

import os
import asyncio
import logging
import re
from typing import Optional

logger = logging.getLogger(__name__)

POLYGON_BASE = "https://api.polygon.io"

# Common English words/abbreviations that match A-Z{1-5} but aren't tickers
_TICKER_BLACKLIST = {
    "A", "I", "IT", "ARE", "ON", "BE", "OR", "AT", "FOR", "IN", "BY", "TO",
    "THE", "AND", "BUT", "NOT", "ALL", "SEC", "IPO", "CEO", "CFO", "COO",
    "YOY", "QOQ", "TTM", "NTM", "LTM", "FY", "Q1", "Q2", "Q3", "Q4",
    "US", "UK", "EU", "GDP", "CPI", "EPS", "PE", "PEG", "DCF", "IRR", "NPV",
    "EBIT", "SGA", "RD", "NA", "OK", "MY", "AM", "PM", "DO", "GO",
    "EV", "ARR", "MRR", "YTD", "MTD", "WTD",
}

# Intent keywords that signal the user wants market/valuation data
_PRICE_INTENT_KEYWORDS = [
    "comp", "comps", "trading multiple", "trading multiples",
    "ev/revenue", "ev/ebitda", "market cap", "mkt cap",
    "price", "share price", "stock price", "valuation",
    "tearsheet", "tear sheet", "peer", "multiples",
    "enterprise value", "build me a comp", "comp set",
    "comp table", "comp sheet", "peer comp",
]


def _looks_like_ticker(s: str) -> bool:
    """1-5 uppercase letters, not a blacklisted common word."""
    return (
        bool(re.match(r'^[A-Z]{1,5}$', s))
        and s not in _TICKER_BLACKLIST
        and len(s) >= 2  # single letters are almost never tickers the user means
    )


def extract_tickers(message: str) -> list[str]:
    """
    Extract stock tickers from a message.
    Matches bare uppercase words (DDOG, SNOW) and $-prefixed ($AAPL).
    De-duped, max 15, filtered against common-word blacklist.
    """
    # $TICKER pattern (explicit)
    dollar_tickers = re.findall(r'\$([A-Z]{1,5})\b', message)
    # Bare uppercase words (only in messages with obvious price intent)
    bare_tickers = re.findall(r'\b([A-Z]{1,5})\b', message)

    seen: set[str] = set()
    result: list[str] = []

    # $-prefixed always win
    for t in dollar_tickers:
        if t not in seen and _looks_like_ticker(t):
            seen.add(t)
            result.append(t)

    # Bare only if message has explicit price intent
    if has_price_intent(message):
        for t in bare_tickers:
            if t not in seen and _looks_like_ticker(t):
                seen.add(t)
                result.append(t)

    return result[:15]


def has_price_intent(message: str) -> bool:
    """True if the message is asking for comp/market data."""
    low = message.lower()
    return any(k in low for k in _PRICE_INTENT_KEYWORDS)


async def get_stock_data(ticker: str) -> dict:
    """
    Fetch previous close price + market cap for one ticker.

    Returns:
        {ticker, price, shares_outstanding_M, market_cap_B, name}
        or {ticker, error} on failure.
    """
    api_key = os.getenv("POLYGON_API_KEY", "")
    if not api_key:
        return {"ticker": ticker, "error": "POLYGON_API_KEY not set"}

    try:
        import httpx
        async with httpx.AsyncClient(timeout=10.0) as client:

            # 1. Previous close (works on free tier)
            prev_resp = await client.get(
                f"{POLYGON_BASE}/v2/aggs/ticker/{ticker}/prev",
                params={"adjusted": "true", "apiKey": api_key},
            )
            prev_data = prev_resp.json()

            if not prev_data.get("results"):
                return {"ticker": ticker, "error": "No price data"}

            close_price = prev_data["results"][0]["c"]

            # 2. Reference data — shares outstanding + company name
            ref_resp = await client.get(
                f"{POLYGON_BASE}/v3/reference/tickers/{ticker}",
                params={"apiKey": api_key},
            )
            ref_data = ref_resp.json()
            ref = ref_data.get("results", {})

            shares = (
                ref.get("weighted_shares_outstanding")
                or ref.get("share_class_shares_outstanding")
                or 0
            )

            market_cap_b = round((close_price * shares) / 1e9, 2) if shares else None
            shares_m = round(shares / 1e6, 1) if shares else None

            return {
                "ticker": ticker,
                "price": round(close_price, 2),
                "shares_outstanding_M": shares_m,
                "market_cap_B": market_cap_b,
                "name": ref.get("name", ticker),
            }

    except Exception as e:
        logger.warning(f"[polygon] {ticker} fetch failed: {e}")
        return {"ticker": ticker, "error": str(e)}


async def get_stocks_batch(tickers: list[str]) -> list[dict]:
    """Fetch multiple tickers concurrently."""
    if not tickers:
        return []
    results = await asyncio.gather(
        *[get_stock_data(t) for t in tickers],
        return_exceptions=True,
    )
    out = []
    for i, r in enumerate(results):
        if isinstance(r, Exception):
            out.append({"ticker": tickers[i], "error": str(r)})
        else:
            out.append(r)
    return out


def format_for_context(stocks: list[dict]) -> str:
    """
    Format stock data as a context block to inject into Claude's prompt.
    Returns empty string if all lookups failed (don't inject noise).
    """
    successes = [s for s in stocks if not s.get("error")]
    if not successes:
        return ""

    lines = ["[MARKET DATA — auto-fetched via Polygon.io (previous close)]"]
    for s in stocks:
        if s.get("error"):
            lines.append(f"  {s.get('ticker', '?')}: data unavailable — {s['error']}")
        else:
            mc = f"${s['market_cap_B']}B" if s.get("market_cap_B") else "N/A"
            sh = f" | {s['shares_outstanding_M']}M shares" if s.get("shares_outstanding_M") else ""
            lines.append(
                f"  {s['ticker']} ({s.get('name', '')}): "
                f"Price ${s['price']}{sh} | Mkt Cap {mc}"
            )

    lines.append(
        "Use these prices and market caps directly in trading multiples — "
        "no need to ask the user for share prices."
    )
    return "\n".join(lines)
