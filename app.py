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


def fetch_revenue_growth_estimates(ticker_obj):
    """Fetches from t.revenue_estimate - Yahoo's actual REVENUE-specific
    consensus growth table, separate from the EPS-focused one above. This
    is what should genuinely drive a DCF's revenue projection, since EPS
    growth and revenue growth aren't the same thing (margin changes, share
    buybacks etc. mean they can diverge)."""
    growth_y1 = growth_y2 = None
    try:
        analysis = ticker_obj.revenue_estimate
    except Exception:
        analysis = None

    if analysis is not None and not analysis.empty and "growth" in analysis.columns:
        if "0y" in analysis.index:
            growth_y1 = safe_float(analysis.loc["0y"].get("growth"))
        if "+1y" in analysis.index:
            growth_y2 = safe_float(analysis.loc["+1y"].get("growth"))
    return growth_y1, growth_y2


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


def project_income_statement(revenue_last, ebit_margin, cogs_pct, sga_pct, rd_pct, da_pct,
                              capex_pct, nwc_pct, tax_rate, growth_path):
    """Same corrected approach as project_fcf: EBIT margin is held flat at
    the real historical margin, not reconstructed from cost-statement lines.
    COGS/SG&A/R&D are still shown for context (scaled at their own
    historical percentage of revenue), but they're informational only here -
    they don't drive EBIT, matching how the Excel template's EBIT line is
    also sourced directly rather than derived from these three."""
    rows = []
    revenue = revenue_last
    for g in growth_path:
        revenue = revenue * (1 + g)
        cogs = revenue * cogs_pct
        gross_profit = revenue - cogs
        sga = revenue * sga_pct
        rd = revenue * rd_pct
        da = revenue * da_pct
        ebit = revenue * ebit_margin
        ebitda = ebit + da
        tax = ebit * tax_rate
        nopat = ebit - tax
        capex = revenue * capex_pct
        nwc_change = revenue * nwc_pct
        fcf = nopat + da - capex - nwc_change
        rows.append(dict(revenue=revenue, growth=g, cogs=cogs, gross_profit=gross_profit,
                          sga=sga, rd=rd, ebitda=ebitda, da=da, ebit=ebit, tax=tax,
                          nopat=nopat, capex=capex, nwc_change=nwc_change, fcf=fcf))
    return rows


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
    """5-year daily P/E, P/S, P/B history - same approach as the Excel
    template's Valuation History tab: a daily price line stepped against
    each fiscal year's actual reported EPS / revenue-per-share /
    book-value-per-share (not a smoothed or interpolated series - each day
    uses whichever fiscal year had most recently reported as of that date).
    Returns a DataFrame with columns date, pe, ps, pb - rows where the
    relevant fundamental wasn't available are simply left out, not guessed.
    """
    try:
        hist = ticker_obj.history(period="5y")
    except Exception:
        return pd.DataFrame(columns=["date", "pe", "ps", "pb"])

    if hist is None or hist.empty or "Close" not in hist.columns:
        return pd.DataFrame(columns=["date", "pe", "ps", "pb"])

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
        return pd.DataFrame(columns=["date", "pe", "ps", "pb"])

    rows = []
    hist_index = hist.index.tz_localize(None) if hist.index.tz is not None else hist.index
    for date, price in zip(hist_index, hist["Close"]):
        # find the most recent fiscal year-end at or before this date
        applicable = None
        for fy_end, eps, rps, bps in fy_data:
            if fy_end <= date:
                applicable = (eps, rps, bps)
            else:
                break
        if applicable is None:
            continue
        eps, rps, bps = applicable
        pe = (price / eps) if (eps and eps > 0) else None
        ps = (price / rps) if (rps and rps > 0) else None
        pb = (price / bps) if (bps and bps > 0) else None
        rows.append(dict(date=date, pe=pe, ps=ps, pb=pb))

    return pd.DataFrame(rows)


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
            forward_pe = safe_float(info.get("forwardPE"))
            forward_revenue = fetch_forward_revenue(t)
            forward_ps = (price * shares / forward_revenue) if (forward_revenue and shares) else None
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
                    # EBIT margin anchored to the correctly-sourced actual EBIT (pulled
                    # directly, not reconstructed from COGS/SG&A/R&D/D&A) - see the
                    # docstring on project_fcf for why this matters.
                    ebit_margin = ebit_last / revenue_last

                    rev_growth_y1, rev_growth_y2 = fetch_revenue_growth_estimates(t)
                    eps_growth_y1, eps_growth_y2, eps_growth_5y = fetch_eps_growth_estimates(t)
                    rev_growth_trailing = safe_float(info.get("revenueGrowth"))

                    # The DCF's actual growth driver: genuine revenue growth where
                    # available, trailing revenue growth as a last-resort proxy -
                    # never EPS growth, which is a different thing (see the two
                    # fetch functions' docstrings above).
                    y1 = rev_growth_y1 if rev_growth_y1 is not None else (rev_growth_trailing or 0.05)
                    y1_source = "Analyst" if rev_growth_y1 is not None else "Trailing (proxy)"

                    st.session_state.results = dict(
                        ticker=ticker_input, company_name=company_name, price=price,
                        shares=shares / 1e6, debt=debt / 1e6, cash=cash / 1e6, beta=beta,
                        target_price=target_price, n_analysts=n_analysts, forward_pe=forward_pe,
                        forward_ps=forward_ps,
                        revenue_last=revenue_last / 1e6, cogs_pct=cogs_pct, sga_pct=sga_pct,
                        rd_pct=rd_pct, da_pct=da_pct, capex_pct=capex_pct, nwc_pct=nwc_pct,
                        tax_rate=tax_rate, ebit_margin=ebit_margin, growth_y1=y1, y1_source=y1_source,
                        rev_growth_y1=rev_growth_y1, rev_growth_y2=rev_growth_y2,
                        eps_growth_y1=eps_growth_y1, eps_growth_y2=eps_growth_y2,
                        # actual historical dollar figures, for the income statement display -
                        # everything above this point only stores ratios/percentages
                        cogs_last=(cogs_last / 1e6) if cogs_last else None,
                        sga_last=(sga_last / 1e6) if sga_last else None,
                        rd_last=(rd_last / 1e6) if rd_last else None,
                        ebit_last=ebit_last / 1e6,
                        da_last=(da_last / 1e6) if da_last else None,
                        last_fiscal_year=str(last_col)[:4],
                    )

                    # Valuation history - fetched once here and stored as a
                    # ready DataFrame, rather than re-fetching on every rerun
                    try:
                        balance_sheet = t.balance_sheet
                    except Exception:
                        balance_sheet = None
                    val_hist_df = fetch_valuation_history(t, income_stmt, balance_sheet)
                    st.session_state.results["val_hist_df"] = val_hist_df
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
        f"Year-1 revenue growth source: **{r['y1_source']}** ({r['growth_y1']:.1%}). "
        f"WACC source: **{wacc_source}** ({wacc:.2%})."
    )

    growth_path = []
    for i in range(5):
        if i == 0:
            growth_path.append(r["growth_y1"])
        elif i == 1 and r["rev_growth_y2"] is not None:
            growth_path.append(r["rev_growth_y2"])
        else:
            # linear taper toward terminal growth - same approach as the Excel template
            anchor = r["rev_growth_y2"] if r["rev_growth_y2"] is not None else r["growth_y1"]
            step = (terminal_growth - anchor) / (5 - 1)
            growth_path.append(anchor + step * i)

    fcf_list = project_fcf(r["revenue_last"], r["ebit_margin"], r["da_pct"], r["capex_pct"],
                            r["nwc_pct"], r["tax_rate"], growth_path)
    implied_price = dcf_implied_price(fcf_list, wacc, terminal_growth, net_debt, r["shares"])

    def metric_card(label, value, delta=None, delta_color="#2D6A4F"):
        """Fully custom metric card, built as raw HTML with hardcoded colors -
        no dependency on Streamlit's internal st.metric() DOM structure or
        class names, which can vary between Streamlit versions and has
        broken twice already relying on CSS targeting those internals."""
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

    st.subheader("Result")
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        metric_card("DCF Implied Price", f"${implied_price:,.2f}" if implied_price else "n/a")
    with m2:
        metric_card("Current Price", f"${r['price']:,.2f}")
    with m3:
        if implied_price:
            upside = (implied_price - r["price"]) / r["price"]
            metric_card("vs. Current", f"{upside:+.1%}")
    with m4:
        if r["target_price"]:
            metric_card("Analyst Target", f"${r['target_price']:,.2f}",
                        f"{r['n_analysts']} analysts" if r["n_analysts"] else None)

    st.write("")  # small spacer
    metric_card("WACC used", f"{wacc:.2%}", wacc_source)

    # ------------------------------------------------------------------
    # Consensus growth - lightweight, just the 4 numbers, not a full
    # projected income statement table. The DCF calculation above still
    # uses the full 5-year projection internally (via project_fcf) - this
    # section just shows what's actually available from analysts, plainly.
    # ------------------------------------------------------------------
    st.subheader("Consensus estimates")

    def fmt_pct_or_na(v):
        return f"{v:+.1%}" if v is not None else "N/A"

    g1, g2 = st.columns(2)
    with g1:
        st.markdown("**Consensus Revenue Growth**")
        st.write(f"Year 1: {fmt_pct_or_na(r['rev_growth_y1'])}")
        st.write(f"Year 2: {fmt_pct_or_na(r['rev_growth_y2'])}")
    with g2:
        st.markdown("**Consensus EPS Growth**")
        st.write(f"Year 1: {fmt_pct_or_na(r['eps_growth_y1'])}")
        st.write(f"Year 2: {fmt_pct_or_na(r['eps_growth_y2'])}")

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
            fcf_g = project_fcf(r["revenue_last"], r["ebit_margin"], r["da_pct"], r["capex_pct"],
                                 r["nwc_pct"], r["tax_rate"],
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
        "The full template additionally supports peer comparison and persistent "
        "manual overrides."
    )

    # ------------------------------------------------------------------
    # Valuation history - same approach as the Excel template's Valuation
    # History tab: daily P/E (and P/S, P/B) against the stock's own past,
    # with mean/+1SD/-1SD stats and a chart showing where today sits.
    # ------------------------------------------------------------------
    st.subheader("Valuation history")
    val_hist_df = r.get("val_hist_df")

    if val_hist_df is None or val_hist_df.empty:
        st.caption("Not enough historical data available for this ticker to build a "
                   "valuation history.")
    else:
        stats = {}
        for metric in ["pe", "ps", "pb"]:
            series = val_hist_df[metric].dropna()
            if len(series) > 1:
                mean = series.mean()
                sd = series.std()
                current = series.iloc[-1]  # most recent trading day's trailing multiple
                stats[metric] = dict(mean=mean, sd=sd, plus1=mean + sd, minus1=mean - sd,
                                      current=current)
            else:
                stats[metric] = dict(mean=None, sd=None, plus1=None, minus1=None, current=None)

        def fmt_x(v):
            return f"{v:.1f}x" if v is not None else "N/A"

        # P/E and P/S "current" rows show forward consensus estimates -
        # P/B stays trailing, since there's no standard "forward book value"
        # concept the way there's a forward earnings or revenue estimate.
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
        }, index=["Mean", "Std Dev", "+1 SD", "-1 SD", "Current (Forward)"])

        st.dataframe(stats_df, width="stretch")
        st.caption("Current (Forward) shows the forward consensus estimate for P/E and "
                   "P/S. P/B stays trailing - there's no standard forward book value "
                   "concept the way there's a forward earnings or revenue estimate.")

        metric_choice = st.selectbox("Chart metric", ["P/E", "P/S", "P/B"])
        metric_key = {"P/E": "pe", "P/S": "ps", "P/B": "pb"}[metric_choice]

        chart_df = val_hist_df[["date", metric_key]].dropna()
        m_stats = stats[metric_key]

        if not chart_df.empty and m_stats["mean"] is not None:
            import altair as alt

            # Combined "Mon YYYY" tick labels - a plain two-line stacked
            # month/year axis isn't reliably renderable without a live
            # browser to verify against, so this achieves the same goal
            # (always knowing which year a point belongs to) with a format
            # that's simple and guaranteed to render correctly.
            x_axis = alt.Axis(title="Date", format="%b %Y", labelAngle=-40, tickCount="month")

            base = alt.Chart(chart_df).mark_line(color="#1F3864").encode(
                x=alt.X("date:T", axis=x_axis),
                y=alt.Y(f"{metric_key}:Q", title=metric_choice),
            )
            mean_line = alt.Chart(pd.DataFrame({"y": [m_stats["mean"]]})).mark_rule(
                color="#5B5F6B", strokeDash=[4, 4]
            ).encode(y="y:Q")
            plus1_line = alt.Chart(pd.DataFrame({"y": [m_stats["plus1"]]})).mark_rule(
                color="#2D6A4F", strokeDash=[2, 2]
            ).encode(y="y:Q")
            minus1_line = alt.Chart(pd.DataFrame({"y": [m_stats["minus1"]]})).mark_rule(
                color="#A13D3D", strokeDash=[2, 2]
            ).encode(y="y:Q")

            layers = base + mean_line + plus1_line + minus1_line
            caption = "Grey dashed = mean. Green dashed = +1 SD. Red dashed = -1 SD."

            # Forward marker - a red dot at today's date, matching the Excel
            # template's chart. Available for P/E and P/S (both have a
            # genuine forward consensus estimate fetched); not for P/B,
            # since there's no standard forward book-value concept.
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
                caption += f" 🔴 = forward {metric_choice} ({forward_value:.1f}x)."

            st.altair_chart(layers, width="stretch")
            st.caption(caption)


st.divider()
st.caption(
    "© 2026 Stock with Claude. This is not financial advice. "
    "Want the full version with peer comparison and valuation history? "
    "[Download the free Excel template](https://youtube.com/@stock_with_claude)."
)
