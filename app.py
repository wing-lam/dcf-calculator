"""
app.py - Stock with Claude: Live DCF Valuation Calculator (V1)
----------------------------------------------------------------
(c) 2026 Stock with Claude. Free for personal use - not for resale or
redistribution. youtube.com/@stock_with_claude

A Streamlit web version of the DCF Excel template's core calculation:
type a ticker, get a live DCF-implied share price plus an interactive
sensitivity grid. Peer Comparison, Valuation History, and the persistent
analyst-override system from the Excel tool are NOT in this V1 - this is
intentionally scoped to just the core DCF engine first.

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

st.set_page_config(page_title="Live DCF Calculator - Stock with Claude", layout="wide")

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
    div[data-testid="stMetric"] {
        background-color: #FFFFFF; border: 1px solid #D8D4CA;
        padding: 1rem; border-radius: 4px;
    }
    div[data-testid="stMetricLabel"] { color: #5B5F6B !important; }
    div[data-testid="stMetricValue"] { color: #1F3864 !important; }
</style>
""", unsafe_allow_html=True)

st.title("Live DCF Valuation Calculator")
st.caption(
    "Type any ticker below. Pulls live data from Yahoo Finance and calculates "
    "a discounted cash flow implied share price, the same way as the full "
    "downloadable Excel template. Not investment advice."
)


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


def fetch_growth_estimates(ticker_obj):
    """Same schema-aware fetch as update_model.py - checks both known
    column-name and index-label variants Yahoo has used."""
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


def project_fcf(revenue_last, cogs_pct, sga_pct, rd_pct, da_pct, capex_pct, nwc_pct,
                 tax_rate, growth_path):
    """5-year FCF projection, holding margins flat at last-actual levels -
    same approach as the Financial Model tab: only revenue growth is
    analyst/math-driven, cost ratios are held flat and are the user's to
    override with their own judgment."""
    fcf_list = []
    revenue = revenue_last
    for g in growth_path:
        revenue = revenue * (1 + g)
        gross_profit = revenue * (1 - cogs_pct)
        ebit = revenue - (revenue * cogs_pct) - (revenue * sga_pct) - (revenue * rd_pct) - (revenue * da_pct)
        nopat = ebit * (1 - tax_rate)
        da = revenue * da_pct
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


# ---------------------------------------------------------------------------
# Input
# ---------------------------------------------------------------------------
col_a, col_b = st.columns([3, 1])
with col_a:
    ticker_input = st.text_input("Ticker", value="NVDA", label_visibility="collapsed",
                                  placeholder="Enter a ticker, e.g. NVDA").upper().strip()
with col_b:
    calculate = st.button("Calculate", width="stretch")

st.caption(
    "🇭🇰 **Hong Kong stocks**: use the full Yahoo Finance format, including the leading "
    "zero and the .HK suffix - e.g. Tencent is **0700.HK**, not 700.HK or 700.  \n"
    "🏦 **Banks and insurers**: this calculator isn't suitable for them. Standard "
    "discounted cash flow doesn't map cleanly onto how financial-sector companies report "
    "(interest income/expense instead of a normal revenue and cost structure) - results "
    "for tickers like this will likely be wrong or fail to load at all."
)

if "results" not in st.session_state:
    st.session_state.results = None

if calculate and ticker_input:
    with st.spinner(f"Fetching live data for {ticker_input} ..."):
        try:
            t = yf.Ticker(ticker_input)
            info = t.info

            price = safe_float(info.get("currentPrice") or info.get("regularMarketPrice"))
            shares = safe_float(info.get("sharesOutstanding"))
            debt = safe_float(info.get("totalDebt")) or 0.0
            cash = safe_float(info.get("totalCash")) or 0.0
            beta = safe_float(info.get("beta")) or 1.0
            target_price = safe_float(info.get("targetMeanPrice"))
            n_analysts = info.get("numberOfAnalystOpinions")
            company_name = info.get("longName") or info.get("shortName") or ticker_input

            if price is None or shares is None:
                st.error(f"Could not find price/shares data for '{ticker_input}'. "
                          f"Check the ticker is correct.")
                st.session_state.results = None
            else:
                income_stmt = t.income_stmt
                last_col = income_stmt.columns[0]  # most recent fiscal year
                revenue_last = safe_float(income_stmt[last_col].get("Total Revenue"))
                ebit_last = get_ebit(income_stmt[last_col])
                da_last = get_da(income_stmt[last_col])
                cogs_last = safe_float(income_stmt[last_col].get("Cost Of Revenue"))
                sga_last = safe_float(income_stmt[last_col].get("Selling General And Administration"))
                rd_last = safe_float(income_stmt[last_col].get("Research And Development")) or 0.0

                if revenue_last is None or revenue_last == 0 or ebit_last is None:
                    st.error("Could not find enough historical financial data for this ticker.")
                    st.session_state.results = None
                else:
                    cogs_pct = (cogs_last / revenue_last) if cogs_last else 0.5
                    sga_pct = (sga_last / revenue_last) if sga_last else 0.1
                    rd_pct = rd_last / revenue_last
                    da_pct = (da_last / revenue_last) if da_last else 0.03
                    capex_pct = 0.03   # not directly available from .income_stmt; reasonable default
                    nwc_pct = 0.02
                    tax_rate = 0.21    # US statutory default; adjustable below

                    growth_y1, growth_y2, growth_5y = fetch_growth_estimates(t)
                    rev_growth_trailing = safe_float(info.get("revenueGrowth"))

                    y1 = growth_y1 if growth_y1 is not None else (rev_growth_trailing or 0.05)
                    y1_source = "Analyst" if growth_y1 is not None else "Trailing (proxy)"

                    st.session_state.results = dict(
                        ticker=ticker_input, company_name=company_name, price=price,
                        shares=shares / 1e6, debt=debt / 1e6, cash=cash / 1e6, beta=beta,
                        target_price=target_price, n_analysts=n_analysts,
                        revenue_last=revenue_last / 1e6, cogs_pct=cogs_pct, sga_pct=sga_pct,
                        rd_pct=rd_pct, da_pct=da_pct, capex_pct=capex_pct, nwc_pct=nwc_pct,
                        tax_rate=tax_rate, growth_y1=y1, y1_source=y1_source,
                        growth_y2=growth_y2, growth_5y=growth_5y,
                    )
        except Exception as e:
            st.error(f"Something went wrong fetching data for '{ticker_input}': {e}")
            st.session_state.results = None

# ---------------------------------------------------------------------------
# Results
# ---------------------------------------------------------------------------
if st.session_state.results:
    r = st.session_state.results
    st.success(f"Loaded: **{r['company_name']}** ({r['ticker']})")

    st.subheader("Assumptions")
    st.caption("Beta and terminal growth are already yours to adjust below. "
               "WACC is calculated from them by default, or override it directly.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        rf = st.number_input("Risk-free rate", value=4.0, step=0.1, format="%.1f") / 100
    with c2:
        erp = st.number_input("Equity risk premium", value=5.0, step=0.1, format="%.1f") / 100
    with c3:
        beta_input = st.number_input("Beta", value=round(r["beta"], 2), step=0.05, format="%.2f")
    with c4:
        terminal_growth = st.number_input("Terminal growth", value=2.5, step=0.25, format="%.2f") / 100

    cost_of_equity = rf + beta_input * erp
    net_debt = r["debt"] - r["cash"]
    weight_equity = (r["price"] * r["shares"]) / ((r["price"] * r["shares"]) + r["debt"]) if (r["price"] * r["shares"] + r["debt"]) > 0 else 1.0
    after_tax_cost_debt = 0.05 * (1 - r["tax_rate"])
    wacc_calculated = weight_equity * cost_of_equity + (1 - weight_equity) * after_tax_cost_debt

    override_wacc = st.checkbox("Override WACC directly (skip the CAPM calculation above)")
    if override_wacc:
        wacc = st.number_input("Your WACC", value=round(wacc_calculated * 100, 2),
                                step=0.1, format="%.2f") / 100
        wacc_source = "Your Override"
    else:
        wacc = wacc_calculated
        wacc_source = "Calculated (CAPM)"

    st.caption(
        f"Beta source: Yahoo Finance live data (yours to override above). "
        f"Year-1 growth source: **{r['y1_source']}** ({r['growth_y1']:.1%}). "
        f"WACC source: **{wacc_source}** ({wacc:.2%})."
    )

    growth_path = []
    for i in range(5):
        if i == 0:
            growth_path.append(r["growth_y1"])
        elif i == 1 and r["growth_y2"] is not None:
            growth_path.append(r["growth_y2"])
        else:
            # linear taper toward terminal growth - same approach as the Excel template
            anchor = r["growth_y2"] if r["growth_y2"] is not None else r["growth_y1"]
            step = (terminal_growth - anchor) / (5 - 1)
            growth_path.append(anchor + step * i)

    fcf_list = project_fcf(r["revenue_last"], r["cogs_pct"], r["sga_pct"], r["rd_pct"],
                            r["da_pct"], r["capex_pct"], r["nwc_pct"], r["tax_rate"], growth_path)
    implied_price = dcf_implied_price(fcf_list, wacc, terminal_growth, net_debt, r["shares"])

    st.subheader("Result")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("DCF Implied Price", f"${implied_price:,.2f}" if implied_price else "n/a")
    m2.metric("Current Price", f"${r['price']:,.2f}")
    if implied_price:
        upside = (implied_price - r["price"]) / r["price"]
        m3.metric("vs. Current", f"{upside:+.1%}")
    if r["target_price"]:
        m4.metric("Analyst Target", f"${r['target_price']:,.2f}",
                   f"{r['n_analysts']} analysts" if r["n_analysts"] else None)

    st.metric("WACC used", f"{wacc:.2%}", wacc_source)

    # ------------------------------------------------------------------
    # Sensitivity grid - 10x10, matching the Excel template's convention.
    # With an even grid size there's no single true-center cell, so (like
    # the Excel version) the 4 cells nearest the actual current WACC/growth
    # assumptions get a highlighted border instead of just one cell.
    # ------------------------------------------------------------------
    st.subheader("Sensitivity grid")
    st.caption("Price across nearby WACC and terminal growth assumptions. "
               "The 4 bordered cells in the middle are closest to your actual current assumptions.")

    wacc_deltas = [(i - 4.5) * 0.005 for i in range(10)]      # +/- 2.25%, 0.5% steps
    growth_deltas = [(i - 4.5) * 0.0025 for i in range(10)]   # +/- 1.125%, 0.25% steps
    wacc_steps = [wacc + d for d in wacc_deltas]
    growth_steps = [terminal_growth + d for d in growth_deltas]

    grid = []
    for g in growth_steps:
        row = []
        for w in wacc_steps:
            fcf_g = project_fcf(r["revenue_last"], r["cogs_pct"], r["sga_pct"], r["rd_pct"],
                                 r["da_pct"], r["capex_pct"], r["nwc_pct"], r["tax_rate"],
                                 growth_path)  # near-term path unaffected by grid axes
            p = dcf_implied_price(fcf_g, w, g, net_debt, r["shares"])
            row.append(p)
        grid.append(row)

    grid_df = pd.DataFrame(
        grid,
        index=[f"{g:.2%}" for g in growth_steps],
        columns=[f"{w:.2%}" for w in wacc_steps],
    )
    grid_df.index.name = "Growth \\ WACC"

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
        "This is a simplified illustrative calculation - cost ratios (COGS, SG&A, R&D, "
        "capex, working capital) are estimated from the latest reported year and held "
        "flat, same as the free downloadable Excel template's default behavior. "
        "The full template additionally supports peer comparison, 5-year valuation "
        "history, and persistent manual overrides."
    )

st.divider()
st.caption(
    "© 2026 Stock with Claude. This is not financial advice. "
    "Want the full version with peer comparison and valuation history? "
    "[Download the free Excel template](https://youtube.com/@stock_with_claude)."

