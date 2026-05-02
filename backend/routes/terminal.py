"""
/terminal — market data endpoints for the tsifl terminal.
Proxies Polygon + yfinance so the browser never needs API keys.
"""

from fastapi import APIRouter, HTTPException
from fastapi.responses import JSONResponse
import os, logging

router = APIRouter()
logger = logging.getLogger(__name__)


@router.get("/quote/{ticker}")
async def terminal_quote(ticker: str):
    """
    Single-ticker quote: price, change, market cap, volume.
    Primary: Polygon. Fallback: yfinance info dict.
    """
    ticker = ticker.upper().strip()
    from services.polygon import get_stock_data

    data = await get_stock_data(ticker)

    # yfinance fallback for extra fields + when Polygon rate-limits
    try:
        import yfinance as yf
        import asyncio
        def _yf():
            t = yf.Ticker(ticker)
            info = t.info or {}
            return info
        info = await asyncio.to_thread(_yf)
    except Exception:
        info = {}

    price      = data.get("price") or float(info.get("currentPrice") or info.get("previousClose") or 0)
    prev_close = float(info.get("previousClose") or info.get("regularMarketPreviousClose") or price)
    change     = round(price - prev_close, 2) if prev_close else 0
    change_pct = round((change / prev_close) * 100, 2) if prev_close else 0

    return {
        "ticker":       ticker,
        "name":         data.get("name") or info.get("longName") or info.get("shortName") or ticker,
        "price":        price,
        "change":       change,
        "change_pct":   change_pct,
        "market_cap_B": data.get("market_cap_B") or round(float(info.get("marketCap") or 0) / 1e9, 2) or None,
        "volume":       info.get("volume") or info.get("regularMarketVolume"),
        "avg_volume":   info.get("averageVolume"),
        "pe":           info.get("trailingPE") or info.get("forwardPE"),
        "eps":          info.get("trailingEps"),
        "week52_high":  info.get("fiftyTwoWeekHigh"),
        "week52_low":   info.get("fiftyTwoWeekLow"),
        "sector":       info.get("sector"),
        "industry":     info.get("industry"),
        "currency":     info.get("currency", "USD"),
    }


@router.get("/fundamentals/{ticker}")
async def terminal_fundamentals(ticker: str):
    """
    Fundamental metrics for the terminal stats panel.
    """
    ticker = ticker.upper().strip()
    from services.yfinance_service import get_fundamentals
    data = await get_fundamentals(ticker)
    if data.get("error"):
        raise HTTPException(404, data["error"])
    return data


@router.get("/history/{ticker}")
async def terminal_history(ticker: str, period: str = "1y", interval: str = "1d"):
    """
    OHLCV history for TradingView chart.
    period: 1d | 5d | 1mo | 3mo | 6mo | 1y | 2y | 5y
    interval: 1m | 5m | 15m | 1h | 1d | 1wk
    """
    ticker = ticker.upper().strip()
    # Whitelist periods/intervals to prevent abuse
    if period not in ("1d","5d","1mo","3mo","6mo","1y","2y","5y","max"):
        period = "1y"
    if interval not in ("1m","5m","15m","1h","1d","1wk","1mo"):
        interval = "1d"
    try:
        import yfinance as yf
        import asyncio

        def _fetch():
            t = yf.Ticker(ticker)
            hist = t.history(period=period, interval=interval)
            if hist.empty:
                return []
            bars = []
            for ts, row in hist.iterrows():
                bars.append({
                    "time":  int(ts.timestamp()),
                    "open":  round(float(row["Open"]),  4),
                    "high":  round(float(row["High"]),  4),
                    "low":   round(float(row["Low"]),   4),
                    "close": round(float(row["Close"]), 4),
                    "volume": int(row["Volume"]) if row["Volume"] == row["Volume"] else 0,
                })
            return bars

        bars = await asyncio.to_thread(_fetch)
        return {"ticker": ticker, "period": period, "interval": interval, "bars": bars}
    except Exception as e:
        raise HTTPException(500, str(e))


@router.get("/search/{query}")
async def terminal_search(query: str):
    """
    Ticker/company name search via yfinance.
    """
    try:
        import yfinance as yf
        import asyncio

        def _search():
            results = yf.Search(query, max_results=8).quotes or []
            out = []
            for r in results:
                out.append({
                    "ticker":   r.get("symbol", ""),
                    "name":     r.get("longname") or r.get("shortname") or r.get("symbol", ""),
                    "exchange": r.get("exchDisp") or r.get("exchange", ""),
                    "type":     r.get("quoteType", ""),
                })
            return out

        results = await asyncio.to_thread(_search)
        return {"results": results}
    except Exception as e:
        return {"results": []}
