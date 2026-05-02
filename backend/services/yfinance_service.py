"""
Yahoo Finance fundamentals service via yfinance.
Free, no API key, quarterly + annual data.
Used as primary fundamentals source (FMP used when available for better data).

Note: Yahoo aggressively rate-limits datacenter IPs (Railway, Heroku, etc.).
We use 24-hour caching + sequential requests + retry to work around this.
"""

import asyncio
import logging
import time
from typing import Optional

logger = logging.getLogger(__name__)

# ── 24-hour cache for fundamentals (quarterly data barely changes) ────────────
_yf_cache: dict[str, tuple[float, dict]] = {}
YF_CACHE_TTL = 86400   # 24 hours — fundamentals are quarterly


def _yf_cache_get(ticker: str) -> dict | None:
    entry = _yf_cache.get(ticker)
    if entry and entry[0] > time.time():
        return entry[1]
    return None


def _yf_cache_set(ticker: str, data: dict):
    _yf_cache[ticker] = (time.time() + YF_CACHE_TTL, data)
    # Evict expired entries if cache grows large
    if len(_yf_cache) > 200:
        now = time.time()
        expired = [k for k, (exp, _) in _yf_cache.items() if exp <= now]
        for k in expired:
            del _yf_cache[k]


def _safe_float(v) -> Optional[float]:
    try:
        return float(v) if v is not None and str(v) != "nan" else None
    except Exception:
        return None


def _to_M(v) -> Optional[float]:
    f = _safe_float(v)
    return round(f / 1e6, 1) if f is not None else None


def _pct(num, den) -> Optional[float]:
    n, d = _safe_float(num), _safe_float(den)
    if n is None or d is None or d == 0:
        return None
    return round((n / d) * 100, 1)


def _growth(curr, prior) -> Optional[float]:
    c, p = _safe_float(curr), _safe_float(prior)
    if c is None or p is None or p == 0:
        return None
    return round(((c - p) / abs(p)) * 100, 1)


def _fetch_sync(ticker: str) -> dict:
    """Synchronous yfinance fetch — run via asyncio.to_thread."""
    try:
        import yfinance as yf
        t = yf.Ticker(ticker)
        info = t.info or {}

        # Try quarterly first, fall back to annual
        try:
            fin = t.quarterly_financials  # rows=metrics, cols=dates
            bs = t.quarterly_balance_sheet
            use_annual = fin is None or fin.empty
        except Exception:
            use_annual = True

        if use_annual:
            try:
                fin = t.financials
                bs = t.balance_sheet
            except Exception:
                fin = None
                bs = None

        def row(df, *keys):
            """Get first matching row from a DataFrame, summing last n cols."""
            if df is None or df.empty:
                return None
            for k in keys:
                matches = [i for i in df.index if k.lower() in str(i).lower()]
                if matches:
                    vals = df.loc[matches[0]].iloc[:4 if not use_annual else 1]
                    total = sum(_safe_float(v) for v in vals if _safe_float(v) is not None)
                    return total if total != 0 else None
            return None

        revenue = row(fin, "Total Revenue")
        gross_profit = row(fin, "Gross Profit")
        op_income = row(fin, "Operating Income", "EBIT")
        ebitda = row(fin, "EBITDA", "Normalized EBITDA")
        net_income = row(fin, "Net Income")

        # Prior year for YoY growth
        revenue_prior = None
        if fin is not None and not fin.empty:
            rev_rows = [i for i in fin.index if "total revenue" in str(i).lower()]
            if rev_rows:
                cols = fin.loc[rev_rows[0]]
                if use_annual and len(cols) >= 2:
                    revenue_prior = _safe_float(cols.iloc[1])
                elif not use_annual and len(cols) >= 8:
                    revenue_prior = sum(
                        _safe_float(v) for v in cols.iloc[4:8]
                        if _safe_float(v) is not None
                    )

        # Balance sheet
        total_debt = None
        cash = None
        if bs is not None and not bs.empty:
            def bs_row(*keys):
                for k in keys:
                    matches = [i for i in bs.index if k.lower() in str(i).lower()]
                    if matches:
                        v = _safe_float(bs.loc[matches[0]].iloc[0])
                        return v
                return None

            short_debt = bs_row("Short Long Term Debt", "Current Debt") or 0
            long_debt = bs_row("Long Term Debt") or 0
            total_debt = short_debt + long_debt
            cash_val = bs_row("Cash And Cash Equivalents", "Cash") or 0
            st_inv = bs_row("Short Term Investments") or 0
            cash = cash_val + st_inv

        net_debt = (total_debt - cash) if total_debt is not None and cash is not None else None

        # ── info-dict fallbacks (more reliable than parsing statement rows) ──
        # yfinance info has ebitda, totalRevenue, grossProfits, netIncomeToCommon
        # as trailing-twelve-month figures. Use them when statement parsing fails.
        if revenue    is None: revenue     = _safe_float(info.get("totalRevenue"))
        if gross_profit is None: gross_profit = _safe_float(info.get("grossProfits"))
        if ebitda     is None: ebitda      = _safe_float(info.get("ebitda"))
        if net_income is None: net_income  = _safe_float(info.get("netIncomeToCommon"))
        if net_debt   is None:
            td = _safe_float(info.get("totalDebt"))
            ca = _safe_float(info.get("totalCash"))
            if td is not None and ca is not None:
                net_debt = td - ca

        # Period label
        period_label = "LTM" if not use_annual else f"FY {info.get('fiscalYearEnd', '')}"
        if fin is not None and not fin.empty:
            latest_col = fin.columns[0]
            year = str(latest_col)[:4]
            period_label = f"FY {year}" if use_annual else f"LTM {year}"

        # Market cap + price from info (fallback when Polygon is rate-limited)
        mc_raw = info.get("marketCap")
        price_raw = info.get("currentPrice") or info.get("previousClose") or info.get("regularMarketPrice")

        return {
            "ticker": ticker,
            "name": info.get("longName") or info.get("shortName") or ticker,
            "period_label": period_label,
            "revenue_ltm_M": _to_M(revenue),
            "gross_margin_pct": _pct(gross_profit, revenue),
            "operating_margin_pct": _pct(op_income, revenue),
            "ebitda_ltm_M": _to_M(ebitda),
            "net_income_ltm_M": _to_M(net_income),
            "revenue_growth_yoy_pct": _growth(revenue, revenue_prior),
            "total_debt_M": _to_M(total_debt),
            "cash_M": _to_M(cash),
            "net_debt_M": _to_M(net_debt),
            # Bonus: price + market cap so templates can fall back if Polygon rate-limits
            "market_cap_B": round(mc_raw / 1e9, 2) if mc_raw else None,
            "price": round(float(price_raw), 2) if price_raw else None,
        }

    except Exception as e:
        logger.warning(f"[yfinance] {ticker} failed: {e}")
        return {"ticker": ticker, "error": str(e)}


async def get_fundamentals(ticker: str) -> dict:
    """Async wrapper — runs yfinance in a thread pool with cache + retry."""
    ticker = ticker.upper()

    # Check 24h cache first
    cached = _yf_cache_get(ticker)
    if cached is not None:
        return cached

    # Retry up to 3 times with backoff (Yahoo rate-limits datacenter IPs)
    last_err = None
    for attempt in range(3):
        if attempt > 0:
            await asyncio.sleep(2.0 * attempt)  # 2s, 4s backoff
        result = await asyncio.to_thread(_fetch_sync, ticker)
        if not result.get("error"):
            _yf_cache_set(ticker, result)
            return result
        last_err = result
        # If rate limited, wait longer
        if "rate" in str(result.get("error", "")).lower():
            await asyncio.sleep(3.0)

    return last_err or {"ticker": ticker, "error": "yfinance failed after retries"}


async def get_fundamentals_batch(tickers: list[str]) -> list[dict]:
    """Fetch fundamentals SEQUENTIALLY (not concurrent) to avoid rate limits."""
    if not tickers:
        return []
    out = []
    for t in tickers:
        # Check cache first — instant if cached
        cached = _yf_cache_get(t.upper())
        if cached is not None:
            out.append(cached)
            continue
        # Sequential with small delay between requests
        result = await get_fundamentals(t)
        out.append(result)
        # Small delay between uncached requests to reduce rate-limit hits
        if not result.get("error"):
            await asyncio.sleep(0.5)
    return out
