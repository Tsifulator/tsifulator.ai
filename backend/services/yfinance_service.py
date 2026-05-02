"""
Yahoo Finance fundamentals service via yfinance.
Free, no API key, quarterly + annual data.
Used as primary fundamentals source (FMP used when available for better data).
"""

import asyncio
import logging
from typing import Optional

logger = logging.getLogger(__name__)


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
    """Async wrapper — runs yfinance in a thread pool."""
    return await asyncio.to_thread(_fetch_sync, ticker)


async def get_fundamentals_batch(tickers: list[str]) -> list[dict]:
    """Fetch fundamentals for multiple tickers concurrently."""
    if not tickers:
        return []
    results = await asyncio.gather(
        *[get_fundamentals(t) for t in tickers],
        return_exceptions=True,
    )
    out = []
    for i, r in enumerate(results):
        if isinstance(r, Exception):
            out.append({"ticker": tickers[i], "error": str(r)})
        else:
            out.append(r)
    return out
