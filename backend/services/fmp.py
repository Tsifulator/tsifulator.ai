"""
Financial Modeling Prep (FMP) service — real fundamentals for comp builders.
Provides LTM income statement + balance sheet data so comps never hallucinate.

Free tier:  limited history
Starter:    $19/mo — full income statement, balance sheet, 5yr history
"""

import os
import asyncio
import logging
from typing import Optional

logger = logging.getLogger(__name__)

FMP_BASE = "https://financialmodelingprep.com/api/v3"


def _api_key() -> str:
    return os.getenv("FMP_API_KEY", "")


async def _get(path: str, params: dict = None) -> dict | list | None:
    """Single FMP GET request. Returns None on any failure."""
    key = _api_key()
    if not key:
        return None
    try:
        import httpx
        p = {"apikey": key, **(params or {})}
        async with httpx.AsyncClient(timeout=10.0) as client:
            r = await client.get(f"{FMP_BASE}{path}", params=p)
            r.raise_for_status()
            return r.json()
    except Exception as e:
        logger.warning(f"[fmp] GET {path} failed: {e}")
        return None


def _sum_quarters(statements: list[dict], field: str, n: int = 4) -> Optional[float]:
    """Sum `field` across the last `n` quarterly statements. Returns None if data missing."""
    vals = []
    for s in statements[:n]:
        v = s.get(field)
        if v is not None:
            vals.append(float(v))
    return sum(vals) if len(vals) == n else (sum(vals) if vals else None)


def _latest(statements: list[dict], field: str) -> Optional[float]:
    """Most recent value of `field`."""
    for s in statements:
        v = s.get(field)
        if v is not None:
            return float(v)
    return None


def _pct(numerator, denominator) -> Optional[float]:
    """Safe percentage: numerator / denominator as 0-100 float."""
    if numerator is None or denominator is None or denominator == 0:
        return None
    return round((numerator / denominator) * 100, 1)


def _growth(current, prior) -> Optional[float]:
    """YoY growth as 0-100 float."""
    if current is None or prior is None or prior == 0:
        return None
    return round(((current - prior) / abs(prior)) * 100, 1)


async def get_fundamentals(ticker: str) -> dict:
    """
    Fetch LTM fundamentals for one ticker.

    Returns dict with:
        ticker, name, period_label,
        revenue_ltm_M, gross_margin_pct, operating_margin_pct, ebitda_ltm_M,
        net_income_ltm_M, revenue_growth_yoy_pct,
        total_debt_M, cash_M, net_debt_M
    On failure returns {ticker, error}.
    """
    if not _api_key():
        return {"ticker": ticker, "error": "FMP_API_KEY not set"}

    # Try quarterly first (paid), fall back to annual (free tier)
    income_data = await _get(
        f"/income-statement/{ticker}",
        {"period": "quarter", "limit": 8},
    )
    use_annual = False
    if not income_data:
        income_data = await _get(
            f"/income-statement/{ticker}",
            {"limit": 2},  # annual, last 2 years
        )
        use_annual = True

    # Balance sheet — quarterly first, fall back to annual
    balance_data = await _get(
        f"/balance-sheet-statement/{ticker}",
        {"period": "quarter", "limit": 2},
    )
    if not balance_data:
        balance_data = await _get(
            f"/balance-sheet-statement/{ticker}",
            {"limit": 1},
        )

    # Company profile for name
    profile_data = await _get(f"/profile/{ticker}")

    if not income_data:
        return {"ticker": ticker, "error": "No income statement data"}

    stmts = income_data if isinstance(income_data, list) else []
    bal = balance_data if isinstance(balance_data, list) else []
    profile = (profile_data[0] if isinstance(profile_data, list) and profile_data else {})

    # LTM: sum last 4 quarters (paid) or use most recent annual (free)
    n = 1 if use_annual else 4
    rev_ltm = _sum_quarters(stmts, "revenue", n)
    gp_ltm = _sum_quarters(stmts, "grossProfit", n)
    op_ltm = _sum_quarters(stmts, "operatingIncome", n)
    ebitda_ltm = _sum_quarters(stmts, "ebitda", n)
    ni_ltm = _sum_quarters(stmts, "netIncome", n)

    # Prior year revenue for YoY growth
    if use_annual:
        rev_prior = _sum_quarters(stmts[1:], "revenue", 1) if len(stmts) >= 2 else None
    else:
        rev_prior = _sum_quarters(stmts[4:], "revenue", 4) if len(stmts) >= 8 else None

    # Period label
    if use_annual:
        period_label = "FY " + str(stmts[0].get("calendarYear", "")) if stmts else "FY"
    else:
        period_label = stmts[0].get("period", "") + " " + str(stmts[0].get("calendarYear", "")) if stmts else "LTM"

    # Balance sheet — most recent quarter
    total_debt = None
    cash = None
    net_debt = None
    if bal:
        b = bal[0]
        short_debt = b.get("shortTermDebt") or 0
        long_debt = b.get("longTermDebt") or 0
        total_debt = float(short_debt) + float(long_debt)
        cash_val = b.get("cashAndCashEquivalents") or 0
        st_inv = b.get("shortTermInvestments") or 0
        cash = float(cash_val) + float(st_inv)
        net_debt = total_debt - cash

    def to_M(v):
        """Convert raw dollars to $M, rounded to 1 decimal."""
        return round(v / 1e6, 1) if v is not None else None

    return {
        "ticker": ticker,
        "name": profile.get("companyName", ticker),
        "period_label": period_label,
        "revenue_ltm_M": to_M(rev_ltm),
        "gross_margin_pct": _pct(gp_ltm, rev_ltm),
        "operating_margin_pct": _pct(op_ltm, rev_ltm),
        "ebitda_ltm_M": to_M(ebitda_ltm),
        "net_income_ltm_M": to_M(ni_ltm),
        "revenue_growth_yoy_pct": _growth(rev_ltm, rev_prior),
        "total_debt_M": to_M(total_debt),
        "cash_M": to_M(cash),
        "net_debt_M": to_M(net_debt),
    }


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


def format_fundamentals_for_context(
    fundamentals: list[dict],
    market_data: list[dict] | None = None,
) -> str:
    """
    Format fundamentals (+ optional market data) into a Claude context block.
    Returns empty string if all lookups failed.
    """
    successes = [f for f in fundamentals if not f.get("error")]
    if not successes:
        return ""

    # Build market data lookup by ticker
    mkt = {}
    for m in (market_data or []):
        if m.get("ticker") and not m.get("error"):
            mkt[m["ticker"]] = m

    lines = [
        "[FUNDAMENTALS — auto-fetched via Financial Modeling Prep (LTM)]",
        "Use ONLY these numbers. Do NOT invent or substitute any values.",
        "",
    ]

    for f in fundamentals:
        if f.get("error"):
            lines.append(f"  {f.get('ticker', '?')}: data unavailable — {f['error']}")
            continue

        t = f["ticker"]
        m = mkt.get(t, {})

        # Market data
        price = m.get("price")
        mkt_cap_b = m.get("market_cap_B")

        # Compute EV if we have both market cap and net debt
        ev_b = None
        net_debt_m = f.get("net_debt_M")
        if mkt_cap_b is not None and net_debt_m is not None:
            ev_b = round(mkt_cap_b + net_debt_m / 1e3, 2)  # net_debt_M → B

        # EV/Revenue
        ev_rev = None
        rev_m = f.get("revenue_ltm_M")
        if ev_b is not None and rev_m:
            ev_rev = round(ev_b * 1e3 / rev_m, 1)  # EV in $B, rev in $M

        def fmt(v, decimals=1, suffix=""):
            return f"{v:,.{decimals}f}{suffix}" if v is not None else "N/A"

        lines.append(f"  {t} — {f.get('name', t)} ({f.get('period_label', 'LTM')})")
        lines.append(f"    Revenue LTM:       ${fmt(rev_m)}M")
        lines.append(f"    Rev Growth YoY:    {fmt(f.get('revenue_growth_yoy_pct'))}%")
        lines.append(f"    Gross Margin:      {fmt(f.get('gross_margin_pct'))}%")
        lines.append(f"    Operating Margin:  {fmt(f.get('operating_margin_pct'))}%")
        lines.append(f"    EBITDA LTM:        ${fmt(f.get('ebitda_ltm_M'))}M")
        lines.append(f"    Net Debt:          ${fmt(net_debt_m)}M")
        if price is not None:
            lines.append(f"    Share Price:       ${fmt(price, 2)}")
        if mkt_cap_b is not None:
            lines.append(f"    Mkt Cap:           ${fmt(mkt_cap_b)}B")
        if ev_b is not None:
            lines.append(f"    EV (calc):         ${fmt(ev_b)}B")
        if ev_rev is not None:
            lines.append(f"    EV/Revenue:        {fmt(ev_rev)}x")
        lines.append("")

    lines.append(
        "Write all revenue/EBITDA values in $M. Write share prices and market caps as plain numbers. "
        "Apply set_number_format: '$#,##0.00' on share price column, '#,##0' on revenue/EBITDA columns, "
        "'0.0%' on margin columns, '0.0x' on multiple columns."
    )
    return "\n".join(lines)
