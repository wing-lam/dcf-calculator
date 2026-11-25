"""
app.py - Stock with Claude: Live DCF Valuation Calculator (V1)
----------------------------------------------------------------
(c) 2026 Stock with Claude. Free for personal use - not for resale or
redistribution. youtube.com/@stock_with_claude

A Streamlit web version of the DCF Excel template's core calculation:
type a ticker, see the stock's price/valuation history and key consensus
figures first, then (further down) a full DCF calculation with an
interactive sensitivity grid. Peer Comparison and the persistent
analyst-override system from the Excel tool are NOT in this V1.

Page order is deliberate: most visitors want the quick informational
picture (price, valuation history, 52-week range, consensus estimates)
rather than to build their own DCF - so that comes first, and the DCF
calculator (which requires more input and judgment) comes last.

Reuses the same validated logic as update_model.py:
- EBIT pulled directly from reported financials (not reconstructed from
  Gross Profit - Opex, which double-counts D&A - see update_model.py's
  docstring for the full story).
- Analyst growth data checked against both known yfinance schema variants
  ('stock' and 'stockTrend' columns; '+5y' and 'LTG' index labels) since
  Yahoo's own field names have changed before without notice.
"""

import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
import math
import time

st.set_page_config(page_title="Live DCF Calculator - Stock with Claude", layout="wide")

# ---------------------------------------------------------------------------
# Language: English by default; Traditional Chinese with ?lang=zh in the link
# (the Chinese Stock Tools page embeds that link) or via the switch below.
# ---------------------------------------------------------------------------
_qp_lang = str(st.query_params.get("lang", "en")).lower()
_lang_choice = st.radio("Language / 語言", ["English", "中文"], index=1 if _qp_lang.startswith("zh") else 0,
                        horizontal=True, label_visibility="collapsed")
ZH = _lang_choice == "中文"
if ZH != _qp_lang.startswith("zh"):
    st.query_params["lang"] = "zh" if ZH else "en"


BETA_CAP = 1.8  # starting beta is capped here; Yahoo's figure is shown and users can override


def T(en, zh):
    """Return the Traditional Chinese text when the Chinese version is selected."""
    return zh if ZH else en


SOURCE_ZH = {"Analyst": "分析師共識", "Trailing (proxy)": "過往增長（替代值）",
             "Default (assumed)": "預設值（假設）",
             "Your Override": "自行輸入", "Calculated (CAPM)": "按 CAPM 計算"}


def src_label(x):
    return SOURCE_ZH.get(x, x) if ZH else x

# ---------------------------------------------------------------------------
# Brand styling - matches the Excel template and website (navy / amber)
# ---------------------------------------------------------------------------
st.markdown("""
<style>
    .main { background-color: #FAF9F6; }
    h1, h2, h3 { color: #1F3864; }
    .stButton>button {
        background-color: #1F3864; color: white; font-weight: 500;
        border-radius: 4px; border: none; padding: 0.5rem 1.5rem;
    }
    .stButton>button:hover { background-color: #14213D; color: white; }

    /* Fixes touch-drag on the price/valuation chart: without this, mobile
       browsers compete with the chart's own touch handling to decide
       whether a drag is "scroll the page" or "interact with the chart",
       which causes both janky/delayed crosshair tracking and the page
       shifting under your finger. Targets the actual svg/canvas elements
       Vega-Lite draws into directly, rather than a Streamlit wrapper class
       name that could vary by version. */
    svg, canvas { touch-action: none; }
</style>
""", unsafe_allow_html=True)

st.title(T("Live DCF Valuation Calculator", "即時 DCF 估值計算器"))
st.caption(T(
    "Type any ticker below. Pulls live data from Yahoo Finance - price and valuation "
    "history first, a full DCF calculation further down the page. Not investment advice.",
    "在下方輸入任何股票代號。工具會從 Yahoo Finance 抓取即時數據：先顯示股價及估值歷史，"
    "頁面下方再做一套完整的現金流折現（DCF）估值。並不構成投資建議。"
))


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def safe_float(x):
    try:
        if x is None:
            return None
        v = float(x)
        if math.isnan(v):  # pandas/NaN counts as missing too, not a real value
            return None
        return v
    except (TypeError, ValueError):
        return None


def get_ebit(income_stmt_col):
    """Same fix as update_model.py: pull EBIT directly from the reported
    line rather than reconstructing it, which double-counts D&A."""
    for label in ["Total Operating Income As Reported", "Operating Income", "EBIT"]:
        if label in income_stmt_col.index:
            val = safe_float(income_stmt_col.get(label))
            if val is not None:
                return val
    return None


def get_da(income_stmt_col):
    for label in ["Reconciled Depreciation", "Depreciation And Amortization", "Depreciation"]:
        if label in income_stmt_col.index:
            val = safe_float(income_stmt_col.get(label))
            if val is not None:
                return val
    return None


def fetch_eps_growth_estimates(ticker_obj):
    """Fetches from t.growth_estimates - despite the generic name, this is
    Yahoo's EARNINGS (EPS) growth estimate table, matching Yahoo's own
    'Growth Estimates' page which is EPS-focused, not revenue. Kept the same
    schema-aware fetch as before (checks both known column-name and
    index-label variants Yahoo has used)."""
    growth_y1 = growth_y2 = growth_5y = None
    try:
        analysis = ticker_obj.growth_estimates
    except Exception:
        analysis = None

    if analysis is not None and not analysis.empty:
        stock_col = None
        for candidate in ("stock", "stockTrend"):
            if candidate in analysis.columns:
                stock_col = candidate
                break
        if stock_col is not None:
            if "0y" in analysis.index:
                growth_y1 = safe_float(analysis.loc["0y"].get(stock_col))
            if "+1y" in analysis.index:
                growth_y2 = safe_float(analysis.loc["+1y"].get(stock_col))
            for ltg_label in ("+5y", "LTG"):
                if ltg_label in analysis.index:
                    growth_5y = safe_float(analysis.loc[ltg_label].get(stock_col))
                    break
    return growth_y1, growth_y2, growth_5y


def _growth_from_table(analysis, year_ago_col):
    """Consensus growth for this year (0y) and next year (+1y). Uses Yahoo's
    'growth' column when present; when Yahoo leaves it out (it often does now),
    works it out from the consensus averages instead:
      0y growth  = this year's average estimate / last year's actual - 1
      +1y growth = next year's average estimate / this year's average - 1"""
    g1 = g2 = None
    if analysis is None or getattr(analysis, "empty", True):
        return g1, g2
    cols = analysis.columns
    if "growth" in cols:
        if "0y" in analysis.index:
            g1 = safe_float(analysis.loc["0y"].get("growth"))
        if "+1y" in analysis.index:
            g2 = safe_float(analysis.loc["+1y"].get("growth"))
    if "avg" in cols:
        avg0 = safe_float(analysis.loc["0y"].get("avg")) if "0y" in analysis.index else None
        avg1 = safe_float(analysis.loc["+1y"].get("avg")) if "+1y" in analysis.index else None
        prev = safe_float(analysis.loc["0y"].get(year_ago_col)) if ("0y" in analysis.index and year_ago_col in cols) else None
        if g1 is None and avg0 and prev and prev > 0:
            g1 = avg0 / prev - 1
        if g2 is None and avg0 and avg1 and avg0 > 0:
            g2 = avg1 / avg0 - 1
    return g1, g2


def fetch_revenue_growth_estimates(ticker_obj):
    """Consensus REVENUE growth (t.revenue_estimate) - what should drive a
    DCF's revenue projection."""
    try:
        analysis = ticker_obj.revenue_estimate
    except Exception:
        analysis = None
    return _growth_from_table(analysis, "yearAgoRevenue")


def fetch_eps_growth_from_estimates(ticker_obj):
    """Fallback EPS growth from t.earnings_estimate when the growth table is empty."""
    try:
        analysis = ticker_obj.earnings_estimate
    except Exception:
        analysis = None
    return _growth_from_table(analysis, "yearAgoEps")


def build_growth_path(g1, g2, terminal_growth):
    """5-year revenue growth: year 1 = g1, year 2 = g2 (if known), then a
    straight-line fade to terminal growth by year 5."""
    path = []
    for i in range(5):
        if i == 0:
            path.append(g1)
        elif i == 1 and g2 is not None:
            path.append(g2)
        else:
            anchor = g2 if g2 is not None else g1
            step = (terminal_growth - anchor) / (5 - 1)
            path.append(anchor + step * i)
    return path



# ---------------------------------------------------------------------------
# Yahoo company data (profile, target price, consensus estimates), fetched
# with retries and cached. Yahoo regularly refuses these requests from shared
# cloud servers such as Streamlit Cloud, while price history still loads.
# Previously the app then fell back silently to placeholders (5% growth,
# beta 1.0) and showed them as if they were real. Now: retry, cache good
# answers for 6 hours, fall back to the last good copy, and otherwise tell
# the visitor plainly that the analyst data is unavailable.
# ---------------------------------------------------------------------------
INFO_KEYS = ("currentPrice", "regularMarketPrice", "sharesOutstanding", "totalDebt", "totalCash",
             "beta", "targetMeanPrice", "numberOfAnalystOpinions", "forwardPE", "longName",
             "shortName", "revenueGrowth")


class YahooIncomplete(Exception):
    """Raised inside the cached fetch so a failed Yahoo response is never cached."""


def _info_with_retry(t, tries=3):
    info = {}
    for i in range(tries):
        try:
            info = t.info or {}
        except Exception:
            info = {}
        if info.get("longName") or info.get("shortName") or info.get("beta") is not None:
            return info, True
        if i < tries - 1:
            time.sleep(1.5 * (i + 1))
    return info, False


@st.cache_data(ttl=6 * 3600, show_spinner=False)
def fetch_company_data(ticker):
    t = yf.Ticker(ticker)
    info, info_ok = _info_with_retry(t)
    rev_y1, rev_y2 = fetch_revenue_growth_estimates(t)
    eps_y1, eps_y2, eps_5y = fetch_eps_growth_estimates(t)
    if eps_y1 is None and eps_y2 is None:
        eps_y1, eps_y2 = fetch_eps_growth_from_estimates(t)
    data = dict(
        info={k: info.get(k) for k in INFO_KEYS},
        rev_growth_y1=rev_y1, rev_growth_y2=rev_y2,
        eps_growth_y1=eps_y1, eps_growth_y2=eps_y2, eps_growth_5y=eps_5y,
        forward_revenue=fetch_forward_revenue(t),
        fetched_at=time.time(),
    )
    data["analyst_ok"] = any(v is not None for v in (
        safe_float(info.get("targetMeanPrice")), rev_y1, rev_y2, eps_y1, eps_y2))
    if not info_ok:
        # Yahoo did not answer properly: don't cache this, try again next time.
        raise YahooIncomplete(data)
    return data


@st.cache_resource
def _last_good_store():
    """Last complete answer per ticker, shared across visitors until the app restarts."""
    return {}


def get_company_data(ticker):
    """Returns (data, status). status is "live" (fresh or cached within 6 hours),
    "stale" (Yahoo not answering now; using the last good copy) or
    "missing" (Yahoo not answering and nothing saved)."""
    store = _last_good_store()
    try:
        data = fetch_company_data(ticker)
        store[ticker] = data
        return data, "live"
    except YahooIncomplete as e:
        partial = e.args[0]
    except Exception:
        partial = dict(info={}, rev_growth_y1=None, rev_growth_y2=None, eps_growth_y1=None,
                       eps_growth_y2=None, eps_growth_5y=None, forward_revenue=None,
                       analyst_ok=False, fetched_at=time.time())
    if ticker in store:
        return store[ticker], "stale"
    return partial, "missing"


def implied_wacc(target, fcf_list, terminal_growth, net_debt, shares):
    """Reverse DCF: the WACC at which the DCF value equals `target` (e.g. the
    current price or the analyst consensus target). DCF value falls as WACC
    rises, so a simple bisection finds it. Returns None if no WACC between
    terminal growth + 0.25% and 40% gives that value."""
    if not target or not fcf_list or not shares:
        return None
    lo, hi = terminal_growth + 0.0025, 0.40
    p_lo = dcf_implied_price(fcf_list, lo, terminal_growth, net_debt, shares)
    p_hi = dcf_implied_price(fcf_list, hi, terminal_growth, net_debt, shares)
    if p_lo is None or p_hi is None or not (p_hi <= target <= p_lo):
        return None
    for _ in range(80):
        mid = (lo + hi) / 2
        p = dcf_implied_price(fcf_list, mid, terminal_growth, net_debt, shares)
        if p > target:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def fetch_forward_revenue(ticker_obj):
    """Consensus current-fiscal-year revenue estimate ('0y' avg) from the
    same t.revenue_estimate table - used to compute a genuine forward P/S,
    the same way Yahoo's own forwardPE uses a forward EPS estimate."""
    try:
        analysis = ticker_obj.revenue_estimate
    except Exception:
        return None
    if analysis is not None and not analysis.empty and "avg" in analysis.columns:
        if "0y" in analysis.index:
            return safe_float(analysis.loc["0y"].get("avg"))
    return None


def project_fcf(revenue_last, ebit_margin, da_pct, capex_pct, nwc_pct, tax_rate, growth_path):
    """5-year FCF projection. EBIT margin is held flat at the REAL historical
    margin (ebit_last / revenue_last, where ebit_last is pulled directly from
    Yahoo's reported EBIT, not reconstructed from cost-statement lines) -
    this avoids the exact bug we caught building the Excel template: a
    bottom-up EBIT reconstruction from only COGS/SG&A/R&D/D&A doesn't
    capture every real cost a company has, so it silently diverges from the
    correctly-sourced actual EBIT margin, creating a jump in Year 1 that
    doesn't correspond to any real assumption changing."""
    fcf_list = []
    revenue = revenue_last
    for g in growth_path:
        revenue = revenue * (1 + g)
        ebit = revenue * ebit_margin
        da = revenue * da_pct
        tax = ebit * tax_rate
        nopat = ebit - tax
        capex = revenue * capex_pct
        nwc_change = revenue * nwc_pct
        fcf = nopat + da - capex - nwc_change
        fcf_list.append(fcf)
    return fcf_list


def dcf_implied_price(fcf_list, wacc, terminal_growth, net_debt, shares):
    if wacc <= terminal_growth or shares in (None, 0):
        return None
    pv_sum = 0.0
    for i, fcf in enumerate(fcf_list):
        pv_sum += fcf / ((1 + wacc) ** (i + 1))
    terminal_value = (fcf_list[-1] * (1 + terminal_growth)) / (wacc - terminal_growth)
    pv_terminal = terminal_value / ((1 + wacc) ** len(fcf_list))
    enterprise_value = pv_sum + pv_terminal
    equity_value = enterprise_value - net_debt
    return equity_value / shares


def fetch_valuation_history(ticker_obj, income_stmt, balance_sheet):
    """5-year daily price, P/E, P/S, P/B history - same approach as the
    Excel template's Valuation History tab: a daily price line stepped
    against each fiscal year's actual reported EPS / revenue-per-share /
    book-value-per-share (not smoothed or interpolated - each day uses
    whichever fiscal year had most recently reported as of that date).
    Returns a DataFrame with columns date, price, pe, ps, pb.
    """
    try:
        hist = ticker_obj.history(period="5y")
    except Exception:
        return pd.DataFrame(columns=["date", "price", "pe", "ps", "pb"])

    if hist is None or hist.empty or "Close" not in hist.columns:
        return pd.DataFrame(columns=["date", "price", "pe", "ps", "pb"])

    # Build a lookup of fiscal-year-end date -> (EPS, revenue, book value per share)
    fy_data = []
    for col in income_stmt.columns:
        eps = safe_float(income_stmt[col].get("Diluted EPS"))
        revenue = safe_float(income_stmt[col].get("Total Revenue"))
        shares_out = safe_float(income_stmt[col].get("Diluted Average Shares"))
        book_value = None
        if balance_sheet is not None and col in balance_sheet.columns:
            book_value = safe_float(balance_sheet[col].get("Stockholders Equity"))
        rev_per_share = (revenue / shares_out) if (revenue and shares_out) else None
        book_per_share = (book_value / shares_out) if (book_value and shares_out) else None
        try:
            fy_end = pd.Timestamp(col).tz_localize(None)
        except Exception:
            continue
        fy_data.append((fy_end, eps, rev_per_share, book_per_share))

    fy_data.sort(key=lambda x: x[0])
    if not fy_data:
        return pd.DataFrame(columns=["date", "price", "pe", "ps", "pb"])

    rows = []
    hist_index = hist.index.tz_localize(None) if hist.index.tz is not None else hist.index
    for date, price in zip(hist_index, hist["Close"]):
        applicable = None
        for fy_end, eps, rps, bps in fy_data:
            if fy_end <= date:
                applicable = (eps, rps, bps)
            else:
                break
        eps, rps, bps = applicable if applicable is not None else (None, None, None)
        pe = (price / eps) if (eps and eps > 0) else None
        ps = (price / rps) if (rps and rps > 0) else None
        pb = (price / bps) if (bps and bps > 0) else None
        rows.append(dict(date=date, price=price, pe=pe, ps=ps, pb=pb))

    return pd.DataFrame(rows)


def compute_52wk_high_low(val_hist_df):
    """52-week high/low with the date each occurred - computed from the
    same daily price history already fetched for the valuation history
    chart, not a separate data source."""
    if val_hist_df is None or val_hist_df.empty:
        return None
    cutoff = val_hist_df["date"].max() - pd.Timedelta(days=365)
    window = val_hist_df[val_hist_df["date"] >= cutoff]
    if window.empty:
        return None
    high_row = window.loc[window["price"].idxmax()]
    low_row = window.loc[window["price"].idxmin()]
    return dict(
        high=high_row["price"], high_date=high_row["date"],
        low=low_row["price"], low_date=low_row["date"],
    )


def metric_card(label, value, delta=None, delta_color="#2D6A4F"):
    """Fully custom metric card, built as raw HTML with hardcoded colors -
    no dependency on Streamlit's internal st.metric() DOM structure or
    class names, which can vary between Streamlit versions and has broken
    twice already relying on CSS targeting those internals."""
    delta_html = (
        f'<div style="color:{delta_color}; font-size:14px; margin-top:4px;">{delta}</div>'
        if delta else ""
    )
    st.markdown(f"""
    <div style="background:#FFFFFF; border:1px solid #D8D4CA; border-radius:4px;
                padding:1rem; height:100%;">
        <div style="color:#5B5F6B; font-size:14px; margin-bottom:4px;">{label}</div>
        <div style="color:#1F3864; font-size:28px; font-weight:600;">{value}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------
col_a, col_b = st.columns([3, 1])
with col_a:
    ticker_input = st.text_input(T("Ticker", "股票代號"), value="NVDA", label_visibility="collapsed",
                                  placeholder=T("Enter a ticker, e.g. NVDA", "輸入股票代號，例如 NVDA")).upper().strip()
with col_b:
    calculate = st.button(T("Calculate", "計算"), width="stretch")

st.caption(T(
    "🇭🇰 **Hong Kong stocks**: use the full Yahoo Finance format, including the leading "
    "zero and the .HK suffix - e.g. Tencent is **0700.HK**, not 700.HK or 700.  \n"
    "🏦 No DCF data shown for banks and insurers.",
    "🇭🇰 **港股**：請使用 Yahoo Finance 的完整格式，包括前面的 0 及 .HK 後綴，"
    "例如騰訊是 **0700.HK**，而不是 700.HK 或 700。  \n"
    "🏦 銀行及保險公司不提供 DCF 估值。"
))

if "results" not in st.session_state:
    st.session_state.results = None

if calculate and ticker_input:
    with st.spinner(f"Fetching live data for {ticker_input} ..."):
        try:
            t = yf.Ticker(ticker_input)
            # Company data (profile, target, estimates) comes from the cached,
            # retried fetch above. If Yahoo is refusing it, fall back to
            # fast_info and the balance sheet for the essentials.
            cdata, data_status = get_company_data(ticker_input)
            info = cdata["info"]

            price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"))
            shares = safe_float(info.get("sharesOutstanding"))
            if price is None or shares is None:
                try:
                    fi = t.fast_info
                    if price is None:
                        price = safe_float(fi.get("lastPrice") or fi.get("previousClose"))
                    if shares is None:
                        shares = safe_float(fi.get("shares"))
                except Exception:
                    pass
            if shares is None:
                try:
                    bs = t.balance_sheet
                    for key in ("Ordinary Shares Number", "Share Issued"):
                        if key in bs.index:
                            shares = safe_float(bs.loc[key].dropna().iloc[0])
                            break
                except Exception:
                    pass
            debt = safe_float(info.get("totalDebt")) or 0.0
            cash = safe_float(info.get("totalCash")) or 0.0
            beta_raw = safe_float(info.get("beta"))
            beta = beta_raw if beta_raw is not None else 1.0
            target_price = safe_float(info.get("targetMeanPrice"))
            n_analysts = info.get("numberOfAnalystOpinions")
            forward_pe = safe_float(info.get("forwardPE"))
            forward_revenue = cdata["forward_revenue"]
            forward_ps = (price * shares / forward_revenue) if (forward_revenue and shares) else None
            company_name = info.get("longName") or info.get("shortName") or ticker_input

            if price is None or shares is None:
                st.error(T(f"Could not find price/shares data for '{ticker_input}'. ", f"找不到「{ticker_input}」的股價或股數數據。") + T(
                          "Check the ticker is correct.", "請檢查股票代號是否正確。"))
                st.session_state.results = None
            else:
                # --- Core info: chart, 52-week range, consensus estimates,
                # target price - fetched regardless of company type, since
                # none of it depends on the revenue/EBIT structure that
                # banks and insurers report differently. ---
                rev_growth_y1, rev_growth_y2 = cdata["rev_growth_y1"], cdata["rev_growth_y2"]
                eps_growth_y1, eps_growth_y2 = cdata["eps_growth_y1"], cdata["eps_growth_y2"]

                results = dict(
                    ticker=ticker_input, company_name=company_name, price=price,
                    shares=shares / 1e6, debt=debt / 1e6, cash=cash / 1e6, beta=beta,
                    target_price=target_price, n_analysts=n_analysts, forward_pe=forward_pe,
                    forward_ps=forward_ps,
                    rev_growth_y1=rev_growth_y1, rev_growth_y2=rev_growth_y2,
                    eps_growth_y1=eps_growth_y1, eps_growth_y2=eps_growth_y2,
                    dcf_available=False,  # set True below only if DCF data is genuinely usable
                    data_status=data_status, analyst_ok=cdata["analyst_ok"],
                    fetched_at=cdata["fetched_at"], beta_missing=beta_raw is None,
                )

                try:
                    balance_sheet = t.balance_sheet
                except Exception:
                    balance_sheet = None
                income_stmt = t.income_stmt
                val_hist_df = fetch_valuation_history(t, income_stmt, balance_sheet)
                results["val_hist_df"] = val_hist_df
                results["wk52"] = compute_52wk_high_low(val_hist_df)

                # --- DCF-specific data: EBIT/revenue and the cost-ratio
                # breakdown. Banks and insurers typically don't report a
                # standard revenue/COGS structure (interest income/expense
                # instead), so this commonly comes back empty for them -
                # in that case, everything above still displays, only the
                # DCF calculator section is skipped further down. ---
                try:
                    last_col = income_stmt.columns[0]
                    revenue_last = safe_float(income_stmt[last_col].get("Total Revenue"))
                    ebit_last = get_ebit(income_stmt[last_col])
                    da_last = get_da(income_stmt[last_col])
                    cogs_last = safe_float(income_stmt[last_col].get("Cost Of Revenue"))
                    sga_last = safe_float(income_stmt[last_col].get("Selling General And Administration"))
                    rd_last = safe_float(income_stmt[last_col].get("Research And Development")) or 0.0
                except Exception:
                    revenue_last = ebit_last = da_last = cogs_last = sga_last = rd_last = None

                if revenue_last is not None and revenue_last != 0 and ebit_last is not None:
                    cogs_pct = (cogs_last / revenue_last) if cogs_last else 0.5
                    sga_pct = (sga_last / revenue_last) if sga_last else 0.1
                    rd_pct = (rd_last or 0.0) / revenue_last
                    da_pct = (da_last / revenue_last) if da_last else 0.03
                    capex_pct = 0.03
                    nwc_pct = 0.02
                    tax_rate = 0.21
                    ebit_margin = ebit_last / revenue_last
                    rev_growth_trailing = safe_float(info.get("revenueGrowth"))
                    if rev_growth_trailing is None:
                        # Yahoo profile missing: use last reported year-on-year revenue growth
                        try:
                            prev_rev = safe_float(income_stmt[income_stmt.columns[1]].get("Total Revenue"))
                            if prev_rev:
                                rev_growth_trailing = revenue_last / prev_rev - 1
                        except Exception:
                            pass
                    if rev_growth_y1 is not None:
                        y1, y1_source = rev_growth_y1, "Analyst"
                    elif rev_growth_trailing is not None:
                        y1, y1_source = rev_growth_trailing, "Trailing (proxy)"
                    else:
                        y1, y1_source = 0.05, "Default (assumed)"

                    results.update(dict(
                        dcf_available=True,
                        revenue_last=revenue_last / 1e6, cogs_pct=cogs_pct, sga_pct=sga_pct,
                        rd_pct=rd_pct, da_pct=da_pct, capex_pct=capex_pct, nwc_pct=nwc_pct,
                        tax_rate=tax_rate, ebit_margin=ebit_margin, growth_y1=y1, y1_source=y1_source,
                        cogs_last=(cogs_last / 1e6) if cogs_last else None,
                        sga_last=(sga_last / 1e6) if sga_last else None,
                        rd_last=(rd_last / 1e6) if rd_last else None,
                        ebit_last=ebit_last / 1e6,
                        da_last=(da_last / 1e6) if da_last else None,
                        last_fiscal_year=str(last_col)[:4],
                    ))

                st.session_state.results = results
        except Exception as e:
            st.error(T(f"Something went wrong fetching data for '{ticker_input}': {e}", f"抓取「{ticker_input}」數據時出錯：{e}"))
            st.session_state.results = None

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if st.session_state.results:
    r = st.session_state.results
    st.success(T(f"Loaded: **{r['company_name']}** ({r['ticker']})", f"已載入：**{r['company_name']}**（{r['ticker']}）"))
    _status = r.get("data_status", "live")
    if _status == "missing":
        st.warning(T(
            "Yahoo Finance isn't returning analyst data right now (target price, consensus estimates, beta). "
            "This happens from time to time on shared cloud servers. Prices and history below are live; "
            "please try again in a few minutes for the full picture.",
            "Yahoo Finance 暫時未有提供分析師數據（目標價、共識預測、Beta 係數）。共用雲端伺服器間中會出現這情況。"
            "以下股價及歷史數據仍是即時的；請稍後再試以取得完整資料。"))
    elif _status == "stale":
        _when = time.strftime("%d %b %Y %H:%M UTC", time.gmtime(r.get("fetched_at", time.time())))
        st.info(T(f"Yahoo Finance isn't responding right now, so analyst data is from the last successful update ({_when}).",
                  f"Yahoo Finance 暫時沒有回應，分析師數據來自上一次成功更新（{_when}）。"))

    val_hist_df = r.get("val_hist_df")

    # ------------------------------------------------------------------
    # 1. Price / valuation chart - defaults to Price, 1 Year
    # ------------------------------------------------------------------
    st.subheader(T("Ticker Valuation History and Price", "估值歷史及股價"))

    if val_hist_df is None or val_hist_df.empty:
        st.caption(T("Not enough historical data available for this ticker to build a "
                   "price or valuation history.", "這隻股票的歷史數據不足，無法建立股價或估值歷史。"))
    else:
        stats = {}
        for metric in ["pe", "ps", "pb"]:
            series = val_hist_df[metric].dropna()
            if len(series) > 1:
                mean = series.mean()
                sd = series.std()
                current = series.iloc[-1]
                stats[metric] = dict(mean=mean, sd=sd, plus1=mean + sd, minus1=mean - sd,
                                      current=current)
            else:
                stats[metric] = dict(mean=None, sd=None, plus1=None, minus1=None, current=None)

        def fmt_x(v):
            return f"{v:.1f}x" if v is not None else "N/A"

        pe_current_display = fmt_x(r.get("forward_pe"))
        ps_current_display = fmt_x(r.get("forward_ps")) if r.get("forward_ps") is not None \
            else fmt_x(stats["ps"]["current"])

        stats_df = pd.DataFrame({
            "P/E": [fmt_x(stats["pe"]["mean"]), fmt_x(stats["pe"]["sd"]),
                    fmt_x(stats["pe"]["plus1"]), fmt_x(stats["pe"]["minus1"]),
                    pe_current_display],
            "P/S": [fmt_x(stats["ps"]["mean"]), fmt_x(stats["ps"]["sd"]),
                    fmt_x(stats["ps"]["plus1"]), fmt_x(stats["ps"]["minus1"]),
                    ps_current_display],
            "P/B": [fmt_x(stats["pb"]["mean"]), fmt_x(stats["pb"]["sd"]),
                    fmt_x(stats["pb"]["plus1"]), fmt_x(stats["pb"]["minus1"]),
                    fmt_x(stats["pb"]["current"])],
        }, index=[T("Mean", "平均值"), T("Std Dev", "標準差"), T("+1 SD", "+1 標準差"),
                   T("-1 SD", "−1 標準差"), T("Current (Forward)", "現時（預測）")])

        st.dataframe(stats_df, width="stretch")
        st.caption(T("Current (Forward) shows the forward consensus estimate for P/E and "
                   "P/S. P/B stays trailing - there's no standard forward book value "
                   "concept the way there's a forward earnings or revenue estimate.",
                   "「現時（預測）」顯示市盈率（P/E）及市銷率（P/S）的預測共識數字。市賬率（P/B）仍採用過往數字，"
                   "因為賬面值並沒有像盈利或收入那樣的標準預測。"))

        CHART_WIDTH, CHART_HEIGHT = 700, 380  # fixed regardless of metric/time-range chosen

        metric_choice = st.selectbox(T("Chart", "圖表"), ["Price", "P/E", "P/S", "P/B"],
                                     format_func=lambda x: T(x, "股價") if x == "Price" else x)

        if metric_choice == "Price":
            range_choice = st.selectbox(T("Time range", "時間範圍"), ["5Y", "3Y", "1Y", "3M"], index=2,
                                        format_func=lambda x: T(x, {"5Y": "5 年", "3Y": "3 年", "1Y": "1 年", "3M": "3 個月"}[x]))
            days_map = {"5Y": 365 * 5, "3Y": 365 * 3, "1Y": 365, "3M": 90}
            cutoff = val_hist_df["date"].max() - pd.Timedelta(days=days_map[range_choice])
            chart_df = val_hist_df[val_hist_df["date"] >= cutoff][["date", "price"]].dropna()
            metric_key, y_title = "price", T("Price ($)", "股價")
        else:
            metric_key = {"P/E": "pe", "P/S": "ps", "P/B": "pb"}[metric_choice]
            chart_df = val_hist_df[["date", metric_key]].dropna()
            y_title = metric_choice

        if chart_df.empty:
            st.caption(T("Not enough data to draw this chart.", "數據不足，無法繪製圖表。"))
        else:
            import altair as alt

            x_axis = alt.Axis(title=T("Date", "日期"), format="%b %Y", labelAngle=-40, tickCount="month")

            # Crosshair pattern: a nearest-point selection bound to
            # pointermove (not the older mouseover-only event) so it
            # responds the same way to a mouse hover on laptop and a
            # finger drag on phone. A wide, fully invisible point layer
            # spans the whole chart so the touch/hover target is generous
            # rather than requiring pixel-precision on the actual line -
            # this is what actually fixes responsiveness on a small touch
            # screen, not just switching event names.
            nearest = alt.selection_point(
                nearest=True, on="mousemove, touchmove, pointermove", fields=["date"], empty=False
            )

            base = alt.Chart(chart_df).mark_line(color="#1F3864").encode(
                x=alt.X("date:T", axis=x_axis),
                y=alt.Y(f"{metric_key}:Q", title=y_title),
            ).properties(width=CHART_WIDTH, height=CHART_HEIGHT)

            # Invisible wide hit-target layer - this is what the pointer/finger
            # actually interacts with, sized generously so a touch nearby
            # (not pixel-exact) still registers.
            hit_target = alt.Chart(chart_df).mark_point(size=400, opacity=0).encode(
                x="date:T",
            ).add_params(nearest)

            # The vertical crosshair line itself, following the pointer
            crosshair_rule = base.mark_rule(color="#B8B8B8", strokeWidth=1).encode(
                opacity=alt.condition(nearest, alt.value(0.8), alt.value(0)),
            ).transform_filter(nearest)

            # Highlighted dot at the exact point the crosshair is on
            crosshair_point = base.mark_point(size=80, filled=True, color="#1F3864").encode(
                opacity=alt.condition(nearest, alt.value(1), alt.value(0)),
            )

            # Date + value labels that move with the crosshair
            date_label = base.mark_text(align="left", dx=8, dy=-18, fontSize=12,
                                         color="#5B5F6B", fontWeight="bold").encode(
                text=alt.condition(nearest, alt.Text("date:T", format="%b %d, %Y"), alt.value("")),
            )
            value_label = base.mark_text(align="left", dx=8, dy=-4, fontSize=13,
                                          color="#1F3864", fontWeight="bold").encode(
                text=alt.condition(nearest, alt.Text(f"{metric_key}:Q", format=",.2f"), alt.value("")),
            )

            layers = base + hit_target + crosshair_rule + crosshair_point + date_label + value_label
            caption = ""

            if metric_choice != "Price":
                m_stats = stats[metric_key]
                if m_stats["mean"] is not None:
                    mean_line = alt.Chart(pd.DataFrame({"y": [m_stats["mean"]]})).mark_rule(
                        color="#5B5F6B", strokeDash=[4, 4]
                    ).encode(y="y:Q")
                    plus1_line = alt.Chart(pd.DataFrame({"y": [m_stats["plus1"]]})).mark_rule(
                        color="#2D6A4F", strokeDash=[2, 2]
                    ).encode(y="y:Q")
                    minus1_line = alt.Chart(pd.DataFrame({"y": [m_stats["minus1"]]})).mark_rule(
                        color="#A13D3D", strokeDash=[2, 2]
                    ).encode(y="y:Q")
                    layers = layers + mean_line + plus1_line + minus1_line
                    caption = T("Grey dashed = mean. Green dashed = +1 SD. Red dashed = -1 SD.",
                              "灰色虛線 = 平均值；綠色虛線 = +1 標準差；紅色虛線 = −1 標準差。")

                forward_value = None
                if metric_key == "pe" and r.get("forward_pe") is not None:
                    forward_value = r["forward_pe"]
                elif metric_key == "ps" and r.get("forward_ps") is not None:
                    forward_value = r["forward_ps"]

                if forward_value is not None:
                    marker_df = pd.DataFrame({
                        "date": [chart_df["date"].max()],
                        metric_key: [forward_value],
                    })
                    marker = alt.Chart(marker_df).mark_point(
                        color="#CC0000", size=120, filled=True
                    ).encode(x="date:T", y=f"{metric_key}:Q")
                    layers = layers + marker
                    caption += T(f" 🔴 = forward {metric_choice} ({forward_value:.1f}x).",
                                f" 🔴 = 預測 {metric_choice}（{forward_value:.1f} 倍）。")

            # width="content" (not "stretch") so the chart keeps the exact
            # fixed size set above no matter which metric/time-range is
            # selected, rather than stretching to fill the column width.
            st.altair_chart(layers, width="content")
            if caption:
                st.caption(caption)

    # ------------------------------------------------------------------
    # 2. 52-week high / low, with dates
    # ------------------------------------------------------------------
    st.subheader(T("52-week range", "52 週高低位"))
    wk52 = r.get("wk52")
    if wk52 is None:
        st.caption(T("Not enough data to compute a 52-week range for this ticker.", "數據不足，無法計算 52 週高低位。"))
    else:
        w1, w2 = st.columns(2)
        with w1:
            metric_card(T("52-Week High", "52 週高位"), f"${wk52['high']:,.2f}",
                        wk52["high_date"].strftime("%b %d, %Y"), delta_color="#5B5F6B")
        with w2:
            metric_card(T("52-Week Low", "52 週低位"), f"${wk52['low']:,.2f}",
                        wk52["low_date"].strftime("%b %d, %Y"), delta_color="#5B5F6B")

    # ------------------------------------------------------------------
    # 3. Consensus growth estimates
    # ------------------------------------------------------------------
    st.subheader(T("Consensus estimates", "分析師共識預測"))

    def fmt_pct_or_na(v):
        return f"{v:+.1%}" if v is not None else "N/A"

    g1, g2 = st.columns(2)
    with g1:
        st.markdown(T("**Consensus Revenue Growth**", "**共識收入增長**"))
        st.write(T("Year 1", "第一年") + f": {fmt_pct_or_na(r['rev_growth_y1'])}")
        st.write(T("Year 2", "第二年") + f": {fmt_pct_or_na(r['rev_growth_y2'])}")
    with g2:
        st.markdown(T("**Consensus EPS Growth**", "**共識每股盈利增長**"))
        st.write(T("Year 1", "第一年") + f": {fmt_pct_or_na(r['eps_growth_y1'])}")
        st.write(T("Year 2", "第二年") + f": {fmt_pct_or_na(r['eps_growth_y2'])}")

    # ------------------------------------------------------------------
    # 4. Consensus analyst target price (moved out of the DCF Result cards)
    # ------------------------------------------------------------------
    if r["target_price"] and not r.get("dcf_available"):
        metric_card(T("Analyst Consensus Target Price", "分析師共識目標價"), f"${r['target_price']:,.2f}",
                    T(f"{r['n_analysts']} analysts", f"{r['n_analysts']} 位分析師") if r["n_analysts"] else None)

    # ------------------------------------------------------------------
    # 5. Consensus starting point: what the analysts' numbers imply, shown
    # BEFORE the visitor changes any assumption (mirrors the Excel version).
    # Consensus growth for years 1-2, fading to 2.5% terminal growth, then
    # walk backwards to the WACC that makes the DCF equal the target price.
    # ------------------------------------------------------------------
    if r.get("dcf_available") and not r.get("analyst_ok", True):
        st.subheader(T("What the consensus implies", "共識預測隱含甚麼"))
        st.info(T("No analyst target or consensus estimates are available for this stock right now, so the "
                  "reverse-DCF panel is skipped. The calculator below starts from the company's own recent "
                  "growth instead; set your own assumptions.",
                  "目前沒有這隻股票的分析師目標價或共識預測，所以略過反向 DCF 面板。下方計算器改以公司近期的實際增長作起點，請自行設定假設。"))

    if r.get("dcf_available") and r.get("analyst_ok", True):
        st.subheader(T("What the consensus implies", "共識預測隱含甚麼"))
        base_tg = 0.025
        base_path = build_growth_path(r["growth_y1"], r["rev_growth_y2"], base_tg)
        base_fcf = project_fcf(r["revenue_last"], r["ebit_margin"], r["da_pct"], r["capex_pct"],
                               r["nwc_pct"], r["tax_rate"], base_path)
        base_nd = r["debt"] - r["cash"]
        wc_tgt = implied_wacc(r.get("target_price"), base_fcf, base_tg, base_nd, r["shares"])
        wc_mkt = implied_wacc(r["price"], base_fcf, base_tg, base_nd, r["shares"])
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            metric_card(T("Consensus target", "共識目標價"), f"${r['target_price']:,.2f}" if r.get("target_price") else "n/a",
                        T(f"{(r['target_price'] / r['price'] - 1):+.1%} vs price", f"較現價 {(r['target_price'] / r['price'] - 1):+.1%}") if r.get("target_price") else None)
        with k2:
            g_txt = f"{r['growth_y1']:+.1%}"
            if r["rev_growth_y2"] is not None:
                g_txt += f" / {r['rev_growth_y2']:+.1%}"
            metric_card(T("Revenue growth Y1 / Y2", "收入增長 第一年／第二年"), g_txt, src_label(r["y1_source"]))
        with k3:
            metric_card(T("WACC implied by target", "目標價隱含的 WACC"), f"{wc_tgt:.2%}" if wc_tgt is not None else "n/a",
                        T("DCF = consensus target", "DCF = 共識目標價"))
        with k4:
            metric_card(T("WACC implied by price", "現價隱含的 WACC"), f"{wc_mkt:.2%}" if wc_mkt is not None else "n/a",
                        T("DCF = current price", "DCF = 現價"))
        st.caption(T("Starting point before you change anything: consensus revenue growth for years 1-2, "
                   "fading to 2.5% terminal growth, with today's margins. The implied WACC is the discount "
                   "rate at which that forecast is worth exactly the analysts' target (or today's price). "
                   "If it looks too low for the stock's risk, the target is demanding; if it looks high, "
                   "the target is conservative. Change the assumptions below to test your own view.",
                   "這是你修改任何假設前的起點：第一及第二年採用共識收入增長，之後逐步降至 2.5% 永續增長率，利潤率維持現有水平。"
                   "隱含 WACC 是令這個預測剛好等於分析師目標價（或現價）的折現率。如果它相對這隻股票的風險顯得太低，"
                   "目標價就偏進取；如果顯得偏高，目標價就偏保守。你可以在下方修改假設，測試自己的看法。"))

    if r.get("dcf_available"):
        st.divider()

        # ==================================================================
        # Everything below this point is the DCF calculator itself - placed
        # last, since most visitors want the information above rather than
        # to build their own valuation.
        # ==================================================================
        st.header(T("DCF calculator", "DCF 計算器"))

        st.subheader(T("Assumptions", "假設"))
        st.caption(T("Beta and terminal growth are already yours to adjust below. "
                   "WACC is calculated from them by default, or override it directly.",
                   "Beta 係數及永續增長率可在下方自行調整。WACC 預設由它們計算，你也可以直接輸入。"))

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            rf = st.number_input(T("Risk-free rate (%)", "無風險利率（%）"), value=4.0, step=0.1, format="%.1f") / 100
        with c2:
            erp = st.number_input(T("Equity risk premium (%)", "股票風險溢價（%）"), value=5.0, step=0.1, format="%.1f") / 100
        with c3:
            beta_input = st.number_input(T("Beta", "Beta 係數（貝他係數）"), value=round(min(r["beta"], BETA_CAP), 2),
                                         step=0.05, format="%.2f")
        with c4:
            terminal_growth = st.number_input(T("Terminal growth (%)", "永續增長率（%）"), value=2.5, step=0.25, format="%.2f") / 100

        cost_of_equity = rf + beta_input * erp
        net_debt = r["debt"] - r["cash"]
        weight_equity = (r["price"] * r["shares"]) / ((r["price"] * r["shares"]) + r["debt"]) if (r["price"] * r["shares"] + r["debt"]) > 0 else 1.0
        after_tax_cost_debt = 0.05 * (1 - r["tax_rate"])
        wacc_calculated = weight_equity * cost_of_equity + (1 - weight_equity) * after_tax_cost_debt

        override_wacc = st.checkbox(T("Override WACC directly (skip the CAPM calculation above)",
                                       "直接輸入 WACC（略過上面的 CAPM 計算）"))
        if override_wacc:
            wacc = st.number_input(T("Your WACC (%)", "你的 WACC（%）"), value=round(wacc_calculated * 100, 2),
                                    step=0.1, format="%.2f") / 100
            wacc_source = "Your Override"
        else:
            wacc = wacc_calculated
            wacc_source = "Calculated (CAPM)"

        beta_note = (T(f" Yahoo shows {r['beta']:.2f}; the starting value is capped at {BETA_CAP:.1f} because very "
                       f"high betas often overstate long-term risk after a volatile spell.",
                       f"Yahoo 顯示為 {r['beta']:.2f}；由於波動期後極高的 Beta 往往高估長期風險，預設值上限為 {BETA_CAP:.1f}。")
                     if r["beta"] > BETA_CAP else "")
        if r.get("beta_missing"):
            beta_note = T(" Yahoo didn't return a beta for this stock, so 1.00 is only a placeholder: set your own.",
                          "Yahoo 未有提供這隻股票的 Beta，1.00 只是暫用值，請自行設定。")
        st.caption(T(
            f"Beta source: Yahoo Finance live data (yours to override above).{beta_note} "
            f"Year-1 revenue growth source: **{r['y1_source']}** ({r['growth_y1']:.1%}). "
            f"WACC source: **{wacc_source}** ({wacc:.2%}).",
            f"Beta 來源：Yahoo Finance 即時數據（可在上方修改）。{beta_note}"
            f"第一年收入增長來源：**{src_label(r['y1_source'])}**（{r['growth_y1']:.1%}）。"
            f"WACC 來源：**{src_label(wacc_source)}**（{wacc:.2%}）。"
        ))

        growth_path = build_growth_path(r["growth_y1"], r["rev_growth_y2"], terminal_growth)

        fcf_list = project_fcf(r["revenue_last"], r["ebit_margin"], r["da_pct"], r["capex_pct"],
                                r["nwc_pct"], r["tax_rate"], growth_path)
        implied_price = dcf_implied_price(fcf_list, wacc, terminal_growth, net_debt, r["shares"])

        st.subheader(T("Result", "結果"))
        m1, m2, m3 = st.columns(3)
        with m1:
            metric_card(T("DCF Implied Price", "DCF 隱含股價"), f"${implied_price:,.2f}" if implied_price else "n/a")
        with m2:
            metric_card(T("Current Price", "現價"), f"${r['price']:,.2f}")
        with m3:
            if implied_price:
                upside = (implied_price - r["price"]) / r["price"]
                metric_card(T("vs. Current", "相對現價"), f"{upside:+.1%}")

        st.write("")
        w_mkt = implied_wacc(r["price"], fcf_list, terminal_growth, net_debt, r["shares"])
        w_tgt = implied_wacc(r.get("target_price"), fcf_list, terminal_growth, net_debt, r["shares"])
        w1, w2, w3 = st.columns(3)
        with w1:
            metric_card(T("WACC used", "採用的 WACC"), f"{wacc:.2%}", src_label(wacc_source))
        with w2:
            metric_card(T("WACC implied by current price", "現價隱含的 WACC"),
                        f"{w_mkt:.2%}" if w_mkt is not None else "n/a",
                        T(f"Discount rate at which the DCF equals ${r['price']:,.2f}", f"令 DCF 等於 ${r['price']:,.2f} 的折現率"))
        with w3:
            tgt_txt = f"${r['target_price']:,.2f}" if r.get("target_price") else T("the target", "目標價")
            metric_card(T("WACC implied by consensus target", "共識目標價隱含的 WACC"),
                        f"{w_tgt:.2%}" if w_tgt is not None else "n/a",
                        T(f"Discount rate at which the DCF equals {tgt_txt}", f"令 DCF 等於 {tgt_txt} 的折現率"))
        st.caption(T("Reverse DCF: with the same cash-flow forecast, these are the discount rates the market "
                   "price and the analysts' average target imply. A lower implied WACC than yours means the "
                   "price already assumes more than your forecast; a higher one means it assumes less. "
                   "n/a means no WACC between terminal growth and 40% produces that price.",
                   "反向 DCF：在同一套現金流預測下，這些是市場價格及分析師平均目標價所隱含的折現率。隱含 WACC 低於你採用的 WACC，"
                   "代表股價已反映比你的預測更樂觀的假設；高於你的，則代表反映的假設較保守。"
                   "n/a 表示在永續增長率至 40% 之間，沒有任何 WACC 能得出該價格。"))

        # ------------------------------------------------------------------
        # Sensitivity grid - 10x10, matching the Excel template's convention.
        # With an even grid size there's no single true-center cell, so (like
        # the Excel version) the 4 cells nearest the actual current WACC/growth
        # assumptions get a highlighted border instead of just one cell.
        # ------------------------------------------------------------------
        st.subheader(T("Sensitivity grid", "敏感度分析表"))
        st.caption(T("Price across nearby WACC and terminal growth assumptions. "
                   "The 4 bordered cells in the middle are closest to your actual current assumptions.",
                   "不同 WACC 及永續增長率組合下的每股價值。中間四個加框的格最接近你目前的假設。"))

        wacc_deltas = [(i - 4.5) * 0.005 for i in range(10)]
        growth_deltas = [(i - 4.5) * 0.0025 for i in range(10)]
        wacc_steps = [wacc + d for d in wacc_deltas]
        growth_steps = [terminal_growth + d for d in growth_deltas]

        grid = []
        for g in growth_steps:
            row = []
            for w in wacc_steps:
                fcf_g = project_fcf(r["revenue_last"], r["ebit_margin"], r["da_pct"], r["capex_pct"],
                                     r["nwc_pct"], r["tax_rate"], growth_path)
                p = dcf_implied_price(fcf_g, w, g, net_debt, r["shares"])
                row.append(p)
            grid.append(row)

        grid_df = pd.DataFrame(
            grid,
            index=[f"{g:.2%}" for g in growth_steps],
            columns=[f"{w:.2%}" for w in wacc_steps],
        )
        grid_df.index.name = T("Growth \\ WACC", "增長率 \\ WACC")

        def fmt_cell(v):
            return f"${v:,.0f}" if v is not None else "n/a"

        def cell_bg(val, vmin, vmax):
            """Muted, desaturated red-to-green background gradient - self-contained,
            no matplotlib dependency (see earlier note on why background_gradient()
            broke the deployed app). Deliberately desaturated rather than vivid,
            since a vivid red background and orange text share the same dominant
            color channel and produce genuinely poor contrast no matter how dark
            the text gets - verified this against actual WCAG contrast math
            before picking these specific tones, not just by eye."""
            if val is None or vmax == vmin:
                return (240, 240, 240)
            frac = max(0.0, min(1.0, (val - vmin) / (vmax - vmin)))
            if frac < 0.5:
                base, target, t = (228, 188, 188), (225, 210, 165), frac * 2
            else:
                base, target, t = (225, 210, 165), (185, 210, 185), (frac - 0.5) * 2
            return tuple(int(base[i] + (target[i] - base[i]) * t) for i in range(3))

        flat_vals = [v for row in grid for v in row if v is not None]
        vmin, vmax = (min(flat_vals), max(flat_vals)) if flat_vals else (0, 1)

        def style_grid(data):
            """Combined position + value aware styling: background color from
            the value (via cell_bg), orange text throughout (verified >=3:1
            contrast against every point on the gradient above), and a red
            border on the 4 center cells specifically."""
            n_rows, n_cols = data.shape
            center_rows = {n_rows // 2 - 1, n_rows // 2}
            center_cols = {n_cols // 2 - 1, n_cols // 2}
            styles = pd.DataFrame("", index=data.index, columns=data.columns)
            for i in range(n_rows):
                for j in range(n_cols):
                    val = data.iloc[i, j]
                    r_, g_, b_ = cell_bg(val, vmin, vmax)
                    css = f"background-color: rgb({r_},{g_},{b_}); color: #B8410A; font-weight: 600;"
                    if i in center_rows and j in center_cols:
                        css += " border: 3px solid #CC0000;"
                    styles.iloc[i, j] = css
            return styles

        st.dataframe(
            grid_df.style.format(fmt_cell).apply(style_grid, axis=None),
            width="stretch",
        )

        st.caption(
            T("This is a simplified illustrative calculation - cost ratios (COGS, SG&A, R&D, "
            "capex, working capital) are estimated from the latest reported year and held "
            "flat, same as the free downloadable Excel template's default behavior. "
            "The full template additionally supports peer comparison and persistent "
            "manual overrides.",
              "這是簡化的示範計算：成本比率（銷售成本、銷售及行政開支、研發、資本開支、營運資金）按最近一個財政年度估算並維持不變，"
              "與可免費下載的 Excel 範本預設做法相同。完整範本另外支援同業比較及保存手動修改。")
        )
    else:
        st.caption(T("No DCF data shown for banks and insurers - standard discounted cash flow doesn't map cleanly onto how financial-sector companies "
                   "report (interest income/expense instead of a normal revenue and cost structure).",
                   "銀行及保險公司不提供 DCF 估值：它們的財務報表以利息收入及支出為主，並沒有一般企業的收入及成本結構，標準現金流折現並不適用。"))


st.divider()
st.caption(T(
    "© 2026 Stock with Claude. This is not financial advice. "
    "Want the full version with peer comparison and valuation history? "
    "[Download the free Excel template](https://youtube.com/@stock_with_claude).",
    "© 2026 Stock with Claude。本工具並不構成投資建議。想要包括同業比較及估值歷史的完整版？"
    "[下載免費 Excel 範本](https://youtube.com/@stock_with_claude)。"
))
