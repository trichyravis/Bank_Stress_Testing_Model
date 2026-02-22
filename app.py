
"""
╔══════════════════════════════════════════════════════════════════╗
║   THE MOUNTAIN PATH - WORLD OF FINANCE                          ║
║   Bank Stress Testing Lab — Prof. V. Ravichandran               ║
╚══════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from scipy import stats
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Bank Stress Testing Lab | Mountain Path",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─────────────────────────────────────────────────────────────────
# MOUNTAIN PATH DESIGN SYSTEM
# ─────────────────────────────────────────────────────────────────

def hex_to_rgba(hex_color, alpha=0.08):
    """Convert hex colour string to rgba() -- safe for all Plotly versions."""
    h = hex_color.lstrip("#")
    if len(h) == 6:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    else:
        r, g, b = 173, 216, 230
    return f"rgba({r},{g},{b},{alpha:.2f})"

COLORS = {
    "darkblue":   "#003366",
    "midblue":    "#004d80",
    "cardBg":     "#112240",
    "bgDark":     "#0a1628",
    "bgGradient": "linear-gradient(135deg,#1a2332,#243447,#2a3f5f)",
    "gold":       "#FFD700",
    "lightblue":  "#ADD8E6",
    "text":       "#e6f1ff",
    "muted":      "#8892b0",
    "green":      "#28a745",
    "red":        "#dc3545",
    "orange":     "#fd7e14",
    "yellow":     "#ffc107",
}

PLOTLY_TEMPLATE = dict(
    paper_bgcolor="#0a1628",
    plot_bgcolor="#112240",
    font=dict(color="#e6f1ff", family="Source Sans Pro"),
    xaxis=dict(gridcolor="#1e3a5f", linecolor="#003366"),
    yaxis=dict(gridcolor="#1e3a5f", linecolor="#003366"),
)

def apply_css():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700;900&family=Source+Sans+Pro:wght@300;400;600&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Source Sans Pro', sans-serif;
        color: {COLORS['text']};
    }}
    .stApp {{
        background: {COLORS['bgGradient']};
        min-height: 100vh;
    }}
    /* Sidebar */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {COLORS['bgDark']}, {COLORS['darkblue']}) !important;
        border-right: 2px solid {COLORS['gold']};
    }}
    [data-testid="stSidebar"] * {{ color: {COLORS['text']} !important; }}
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stSlider label,
    [data-testid="stSidebar"] .stNumberInput label {{ color: {COLORS['gold']} !important; font-weight: 600; }}

    /* Cards */
    .metric-card {{
        background: {COLORS['cardBg']};
        border: 1px solid {COLORS['midblue']};
        border-top: 3px solid {COLORS['gold']};
        border-radius: 8px;
        padding: 16px 20px;
        margin: 6px 0;
    }}
    .metric-label {{
        color: {COLORS['muted']};
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }}
    .metric-value {{
        color: {COLORS['gold']};
        font-size: 1.8rem;
        font-weight: 700;
        font-family: 'Playfair Display', serif;
    }}
    .metric-delta {{
        font-size: 0.82rem;
        font-weight: 600;
    }}
    .delta-pos {{ color: {COLORS['green']}; }}
    .delta-neg {{ color: {COLORS['red']}; }}
    .delta-warn {{ color: {COLORS['orange']}; }}

    /* Section headers */
    .section-header {{
        font-family: 'Playfair Display', serif;
        font-size: 1.4rem;
        font-weight: 700;
        color: {COLORS['gold']};
        border-bottom: 2px solid {COLORS['midblue']};
        padding-bottom: 8px;
        margin: 18px 0 14px 0;
    }}
    .hero-title {{
        font-family: 'Playfair Display', serif;
        font-size: 2.2rem;
        font-weight: 900;
        color: {COLORS['gold']};
        letter-spacing: 1px;
    }}
    .hero-subtitle {{
        color: {COLORS['lightblue']};
        font-size: 1.0rem;
        font-weight: 300;
    }}
    .brand-tag {{
        color: {COLORS['muted']};
        font-size: 0.75rem;
        letter-spacing: 2px;
        text-transform: uppercase;
    }}

    /* Risk badges */
    .badge-green  {{ background:#0d3320; color:{COLORS['green']}; border:1px solid {COLORS['green']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-yellow {{ background:#332900; color:{COLORS['yellow']}; border:1px solid {COLORS['yellow']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-orange {{ background:#332000; color:{COLORS['orange']}; border:1px solid {COLORS['orange']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}
    .badge-red    {{ background:#330d10; color:{COLORS['red']}; border:1px solid {COLORS['red']}; border-radius:4px; padding:2px 10px; font-size:0.78rem; font-weight:700; }}

    /* Tables */
    .styled-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
    .styled-table th {{ background:{COLORS['darkblue']}; color:{COLORS['gold']}; padding:8px 12px; text-align:left; font-weight:700; letter-spacing:0.5px; }}
    .styled-table td {{ padding:7px 12px; border-bottom:1px solid #1e3a5f; color:{COLORS['text']}; }}
    .styled-table tr:hover td {{ background:#243447; }}

    /* Scenario pill */
    .scenario-pill {{
        display:inline-block;
        background:{COLORS['midblue']};
        color:{COLORS['gold']};
        border-radius:20px;
        padding:3px 14px;
        font-size:0.78rem;
        font-weight:700;
        letter-spacing:0.5px;
        margin:2px;
    }}
    /* Divider */
    .gold-divider {{ border:none; border-top:1px solid {COLORS['gold']}44; margin:16px 0; }}

    /* Override Streamlit defaults */
    .stTabs [data-baseweb="tab-list"] {{ background:{COLORS['cardBg']}; border-radius:8px; padding:4px; }}
    .stTabs [data-baseweb="tab"] {{ color:{COLORS['muted']}; font-weight:600; }}
    .stTabs [aria-selected="true"] {{ background:{COLORS['darkblue']} !important; color:{COLORS['gold']} !important; border-radius:6px; }}
    div[data-testid="stMetricValue"] {{ color:{COLORS['gold']}; font-family:'Playfair Display',serif; }}

    /* ── Selectbox: selected value text ── */
    [data-testid="stSidebar"] [data-baseweb="select"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-baseweb="select"] span,
    [data-testid="stSidebar"] [data-baseweb="select"] div,
    [data-testid="stSidebar"] .stSelectbox div[data-baseweb="select"] > div {{
        color: {COLORS['text']} !important;
        font-weight: 600 !important;
    }}
    /* Selectbox container box */
    [data-testid="stSidebar"] [data-baseweb="select"] > div:first-child {{
        background-color: {COLORS['cardBg']} !important;
        border: 1px solid {COLORS['midblue']} !important;
        border-radius: 6px !important;
    }}
    [data-testid="stSidebar"] [data-baseweb="select"] > div:first-child:hover {{
        border-color: {COLORS['gold']} !important;
    }}
    /* ── Dropdown popup list ── */
    [data-baseweb="popover"] ul,
    [data-baseweb="menu"] ul,
    [role="listbox"] {{
        background-color: {COLORS['cardBg']} !important;
        border: 1px solid {COLORS['midblue']} !important;
    }}
    /* Each option item */
    [data-baseweb="menu"] li,
    [role="option"] {{
        background-color: {COLORS['cardBg']} !important;
        color: {COLORS['text']} !important;
        font-weight: 500 !important;
    }}
    [data-baseweb="menu"] li:hover,
    [role="option"]:hover {{
        background-color: {COLORS['darkblue']} !important;
        color: {COLORS['gold']} !important;
    }}
    /* Selected/highlighted option */
    [aria-selected="true"][role="option"],
    [data-baseweb="menu"] li[aria-selected="true"] {{
        background-color: {COLORS['darkblue']} !important;
        color: {COLORS['gold']} !important;
        font-weight: 700 !important;
    }}
    /* Dropdown chevron icon */
    [data-testid="stSidebar"] [data-baseweb="select"] svg {{
        fill: {COLORS['gold']} !important;
    }}
    </style>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
# SYNTHETIC BANK DATA
# ─────────────────────────────────────────────────────────────────
def generate_bank_data():
    """Generate a realistic synthetic Indian mid-size bank."""
    np.random.seed(42)
    bank = {
        "name": "Mountain Path Bank Ltd.",
        "type": "Private Sector Scheduled Commercial Bank",
        "rating": "AA-",
        # ── Balance Sheet (₹ Crore) ──
        "total_assets": 185000,
        "gross_loans": 108000,
        "retail_loans": 38000,
        "corporate_loans": 45000,
        "msme_loans": 18000,
        "agri_loans": 7000,
        "gross_npa": 7560,       # 7.0% gross NPA
        "net_npa": 3780,
        "provisions": 3780,
        "investments_htm": 28000,
        "investments_afs": 14000,
        "equity_portfolio": 3200,
        "cash_hqla": 22000,
        "fixed_assets": 1800,
        "other_assets": 8000,
        # ── Liabilities ──
        "total_deposits": 148000,
        "casa_deposits": 59200,   # 40% CASA
        "term_deposits": 88800,
        "wholesale_funding": 18500,
        "sub_debt": 3200,
        "equity_capital": 15300,
        # ── Capital ──
        "cet1_capital": 13500,
        "tier1_capital": 15000,
        "tier2_capital": 3200,
        "total_capital": 18200,
        "rwa": 141000,
        # ── Income (Annual, ₹ Crore) ──
        "net_interest_income": 7200,
        "fee_income": 2100,
        "trading_income": 680,
        "other_income": 420,
        "operating_costs": 4800,
        "pre_provision_profit": 5600,
        "provisions_charge": 2800,
        "pbt": 2800,
        "tax": 700,
        "pat": 2100,
        # ── Key Ratios ──
        "nim": 3.9,
        "roe": 13.7,
        "roa": 1.1,
        "casa_ratio": 40.0,
        "pcr": 50.0,
        "lcr": 142.0,
        "nsfr": 118.0,
        "crar": 12.9,
        "cet1_ratio": 9.57,
        "leverage_ratio": 8.1,
        "cd_ratio": 72.9,
        # ── Asset Quality ──
        "gross_npa_ratio": 7.0,
        "net_npa_ratio": 3.5,
        "slma_ratio": 3.2,
        "restructured_ratio": 1.8,
        # ── Sector Exposure ──
        "sector_real_estate": 12000,
        "sector_infrastructure": 18000,
        "sector_textile": 5500,
        "sector_gems": 3200,
        "sector_nbfc": 9500,
        "sector_power": 8800,
        # ── Duration ──
        "avg_duration_assets": 4.8,
        "avg_duration_liabs": 2.1,
        "repricing_gap": 18500,   # rate-sensitive assets - liabilities
    }
    return bank

def generate_macro_baseline():
    """8-quarter baseline macro path."""
    return pd.DataFrame({
        "quarter": [f"Q{i}" for i in range(1, 9)],
        "gdp_growth": [6.8, 7.1, 7.0, 6.9, 7.2, 7.3, 7.1, 7.4],
        "inflation": [4.2, 4.0, 3.9, 4.1, 4.3, 4.2, 4.0, 3.8],
        "repo_rate": [6.50]*8,
        "usd_inr": [83.2, 83.5, 83.8, 84.0, 84.2, 84.5, 84.7, 85.0],
        "nifty_index": [22500, 23100, 23600, 24000, 24500, 25000, 25400, 25900],
        "unemployment": [7.8]*8,
        "property_price_chg": [5.2, 5.5, 5.8, 5.5, 5.2, 5.0, 5.3, 5.6],
        "credit_spread_bps": [180]*8,
    })

# ─────────────────────────────────────────────────────────────────
# SCENARIO LIBRARY
# ─────────────────────────────────────────────────────────────────
SCENARIOS = {
    "🟢 Baseline": {
        "color": COLORS["green"],
        "severity": "Baseline",
        "description": "Business-as-usual with stable macro environment.",
        "gdp_shock": [0.0]*8,
        "rate_shock_bps": 0,
        "equity_shock_pct": 0.0,
        "inr_depreciation": 0.0,
        "npa_multiplier": 1.0,
        "deposit_runoff": 0.0,
        "property_shock": 0.0,
        "credit_spread_widening": 0,
        "fee_income_shock": 0.0,
    },
    "🟡 Mild Stress": {
        "color": COLORS["yellow"],
        "severity": "Mild",
        "description": "Moderate slowdown: RBI rate hike cycle, mild NPA uptick.",
        "gdp_shock": [-0.8, -1.0, -0.9, -0.7, -0.5, -0.4, -0.3, -0.2],
        "rate_shock_bps": 75,
        "equity_shock_pct": -0.12,
        "inr_depreciation": 0.04,
        "npa_multiplier": 1.3,
        "deposit_runoff": 0.02,
        "property_shock": -0.05,
        "credit_spread_widening": 50,
        "fee_income_shock": -0.08,
    },
    "🟠 Moderate Stress": {
        "color": COLORS["orange"],
        "severity": "Moderate",
        "description": "Significant macro downturn: global spillover, corporate NPA surge.",
        "gdp_shock": [-1.5, -2.5, -2.8, -2.2, -1.8, -1.2, -0.8, -0.5],
        "rate_shock_bps": 150,
        "equity_shock_pct": -0.28,
        "inr_depreciation": 0.10,
        "npa_multiplier": 1.8,
        "deposit_runoff": 0.06,
        "property_shock": -0.15,
        "credit_spread_widening": 120,
        "fee_income_shock": -0.18,
    },
    "🔴 Severe Stress": {
        "color": COLORS["red"],
        "severity": "Severe",
        "description": "Deep recession: GDP contraction, crisis-level NPA, equity crash.",
        "gdp_shock": [-2.5, -4.5, -5.2, -4.8, -3.5, -2.2, -1.5, -0.8],
        "rate_shock_bps": 250,
        "equity_shock_pct": -0.45,
        "inr_depreciation": 0.18,
        "npa_multiplier": 2.5,
        "deposit_runoff": 0.12,
        "property_shock": -0.28,
        "credit_spread_widening": 280,
        "fee_income_shock": -0.32,
    },
    "⚫ Extreme / Tail": {
        "color": "#cc66ff",
        "severity": "Extreme",
        "description": "Systemic crisis: GFC-2008 severity applied to Indian banking context.",
        "gdp_shock": [-3.5, -6.0, -7.3, -6.5, -5.0, -3.5, -2.0, -1.0],
        "rate_shock_bps": 350,
        "equity_shock_pct": -0.60,
        "inr_depreciation": 0.28,
        "npa_multiplier": 3.8,
        "deposit_runoff": 0.20,
        "property_shock": -0.40,
        "credit_spread_widening": 500,
        "fee_income_shock": -0.45,
    },
    "🏗️ Real Estate Collapse": {
        "color": "#ff9966",
        "severity": "Sector-specific",
        "description": "Property prices crash 35%: developer NPAs cascade to bank balance sheets.",
        "gdp_shock": [-1.2, -2.0, -2.5, -2.0, -1.5, -1.0, -0.8, -0.5],
        "rate_shock_bps": 50,
        "equity_shock_pct": -0.22,
        "inr_depreciation": 0.06,
        "npa_multiplier": 2.2,
        "deposit_runoff": 0.05,
        "property_shock": -0.35,
        "credit_spread_widening": 150,
        "fee_income_shock": -0.20,
    },
    "📈 Rate Shock": {
        "color": "#66ccff",
        "severity": "Market risk",
        "description": "Sudden 300 bps rate hike: MTM losses on investment portfolio, NII pressure.",
        "gdp_shock": [-0.5, -0.8, -0.7, -0.5, -0.3, -0.2, -0.1, 0.0],
        "rate_shock_bps": 300,
        "equity_shock_pct": -0.18,
        "inr_depreciation": -0.03,
        "npa_multiplier": 1.2,
        "deposit_runoff": 0.04,
        "property_shock": -0.10,
        "credit_spread_widening": 80,
        "fee_income_shock": -0.05,
    },
    "💧 Liquidity Crisis": {
        "color": "#33cccc",
        "severity": "Liquidity",
        "description": "IL&FS-style contagion: wholesale funding freeze, deposit run-off surge.",
        "gdp_shock": [-1.0, -1.5, -1.2, -0.8, -0.5, -0.3, -0.2, -0.1],
        "rate_shock_bps": 100,
        "equity_shock_pct": -0.30,
        "inr_depreciation": 0.08,
        "npa_multiplier": 1.6,
        "deposit_runoff": 0.18,
        "property_shock": -0.12,
        "credit_spread_widening": 220,
        "fee_income_shock": -0.25,
    },
}

# ─────────────────────────────────────────────────────────────────
# STRESS TEST ENGINE
# ─────────────────────────────────────────────────────────────────
def run_stress_test(bank, scenario, custom_params=None):
    """Core stress testing engine — projects 8 quarters."""
    sc = scenario.copy()
    if custom_params:
        sc.update(custom_params)

    quarters = 8
    macro = generate_macro_baseline()

    results = []
    capital = bank["cet1_capital"]
    tier1   = bank["tier1_capital"]
    total_capital = bank["total_capital"]
    rwa     = bank["rwa"]
    gross_npa = bank["gross_npa"]
    gross_loans = bank["gross_loans"]
    deposits = bank["total_deposits"]
    hqla    = bank["cash_hqla"]

    # One-time market shocks (applied at Q1)
    afs_mtm_loss = (bank["avg_duration_assets"] - 2.0) * \
                   (sc["rate_shock_bps"] / 10000) * bank["investments_afs"]
    equity_mtm_loss = abs(sc["equity_shock_pct"]) * bank["equity_portfolio"]
    fx_loss = sc["inr_depreciation"] * (bank["gross_loans"] * 0.04)  # ~4% forex exposure

    for q in range(quarters):
        # ── GDP shock path ──
        gdp_q = macro["gdp_growth"].iloc[q] + sc["gdp_shock"][q]

        # ── Credit Loss Model ──
        gdp_sensitivity = 2.8   # 1% GDP fall → 2.8% relative NPA increase
        gdp_drag = max(-sc["gdp_shock"][q], 0)
        npa_ratio_stressed = min(
            (bank["gross_npa_ratio"] / 100) * sc["npa_multiplier"] *
            (1 + gdp_sensitivity * gdp_drag / 100),
            0.28
        )
        gross_npa_q = gross_loans * npa_ratio_stressed

        # Incremental provisions (60% PCR on fresh NPAs)
        prev_npa = gross_npa if q == 0 else results[-1]["gross_npa"]
        incremental_npa = max(gross_npa_q - prev_npa, 0)
        incremental_provisions = incremental_npa * 0.60

        # ── NII Impact ──
        rate_bps = sc["rate_shock_bps"] if q < 2 else sc["rate_shock_bps"] * 0.7
        nii_rate_impact = bank["repricing_gap"] * (rate_bps / 10000) * 0.25
        nii_quarterly = bank["net_interest_income"] / 4 + nii_rate_impact * (1 if q >= 1 else 0)

        # ── Fee Income ──
        fee_q = bank["fee_income"] / 4 * (1 + sc["fee_income_shock"] * (q/7))

        # ── Trading / MTM (Q1 one-time shock) ──
        mtm_loss_q = (afs_mtm_loss + equity_mtm_loss + fx_loss) if q == 0 else 0
        trading_q = bank["trading_income"] / 4 - mtm_loss_q

        # ── P&L ──
        ppop_q = nii_quarterly + fee_q + trading_q - bank["operating_costs"] / 4
        net_profit_q = ppop_q - incremental_provisions - bank["tax"] / 4

        # ── Capital ──
        capital += net_profit_q
        tier1 += net_profit_q * 0.85
        total_capital += net_profit_q * 0.90
        rwa *= (1 + max(-sc["gdp_shock"][q], 0) * 0.003 + 0.003)  # mild RWA inflation

        cet1_ratio = capital / rwa * 100
        tier1_ratio = tier1 / rwa * 100
        crar = total_capital / rwa * 100

        # ── Liquidity ──
        deposit_runoff_q = sc["deposit_runoff"] * deposits * (1.2 if q == 0 else 0.5)
        deposits = max(deposits - deposit_runoff_q, deposits * 0.6)
        hqla = max(hqla - deposit_runoff_q * 0.3, 0)
        lcr = (hqla / max(deposits * 0.03 + deposit_runoff_q, 1)) * 100

        # ── Property / Collateral ──
        collateral_loss = abs(sc["property_shock"]) * bank["sector_real_estate"] * 0.20 \
                          if q == 0 else 0
        capital -= collateral_loss if q == 0 else 0

        results.append({
            "quarter": f"Q{q+1}",
            "gdp_growth": gdp_q,
            "gross_npa_ratio": npa_ratio_stressed * 100,
            "gross_npa": gross_npa_q,
            "incremental_provisions": incremental_provisions,
            "nii": nii_quarterly,
            "fee_income": fee_q,
            "trading_income": trading_q,
            "ppop": ppop_q,
            "net_profit": net_profit_q,
            "cet1_capital": capital,
            "tier1_capital": tier1,
            "total_capital": total_capital,
            "rwa": rwa,
            "cet1_ratio": cet1_ratio,
            "tier1_ratio": tier1_ratio,
            "crar": crar,
            "lcr": lcr,
            "deposits": deposits,
            "hqla": hqla,
            "mtm_loss": mtm_loss_q,
            "breach_cet1": cet1_ratio < 8.0,
            "breach_tier1": tier1_ratio < 9.5,
            "breach_crar": crar < 11.5,
            "breach_lcr": lcr < 100,
        })

    df = pd.DataFrame(results)
    # Summary stats
    summary = {
        "min_cet1": df["cet1_ratio"].min(),
        "min_crar": df["crar"].min(),
        "min_lcr": df["lcr"].min(),
        "peak_npa": df["gross_npa_ratio"].max(),
        "cumulative_provisions": df["incremental_provisions"].sum(),
        "cumulative_net_profit": df["net_profit"].sum(),
        "total_mtm_loss": df["mtm_loss"].sum(),
        "any_breach": df[["breach_cet1","breach_tier1","breach_crar","breach_lcr"]].any().any(),
        "cet1_breaches": df["breach_cet1"].sum(),
    }
    return df, summary

# ─────────────────────────────────────────────────────────────────
# SENSITIVITY ANALYSIS
# ─────────────────────────────────────────────────────────────────
def run_sensitivity(bank, variable, values, base_scenario):
    """Sweep a single variable and record min CET1."""
    results = []
    for val in values:
        sc = base_scenario.copy()
        sc[variable] = val
        if variable == "rate_shock_bps":
            pass
        elif variable == "npa_multiplier":
            sc["npa_multiplier"] = val
        elif variable == "equity_shock_pct":
            sc["equity_shock_pct"] = -abs(val)
        _, summary = run_stress_test(bank, sc)
        results.append({"value": val, "min_cet1": summary["min_cet1"],
                         "peak_npa": summary["peak_npa"]})
    return pd.DataFrame(results)

# ─────────────────────────────────────────────────────────────────
# PLOTTING HELPERS
# ─────────────────────────────────────────────────────────────────
def base_fig(**kwargs):
    fig = go.Figure(**kwargs)
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        margin=dict(l=40, r=20, t=50, b=40),
        legend=dict(bgcolor="#0a1628", bordercolor="#003366", borderwidth=1,
                    font=dict(color=COLORS["text"])),
    )
    fig.update_xaxes(gridcolor="#1e3a5f", linecolor="#003366", tickfont=dict(color=COLORS["muted"]))
    fig.update_yaxes(gridcolor="#1e3a5f", linecolor="#003366", tickfont=dict(color=COLORS["muted"]))
    return fig

def plot_multi_scenario_cet1(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["cet1_ratio"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
            marker=dict(size=6),
        ))
    # Regulatory floors
    fig.add_hline(y=8.0, line=dict(color="#dc3545", dash="dash", width=1.5),
                  annotation_text="CET1 Min (8%)", annotation_font_color="#dc3545")
    fig.add_hline(y=9.5, line=dict(color="#fd7e14", dash="dot", width=1.2),
                  annotation_text="RBI Buffer (9.5%)", annotation_font_color="#fd7e14")
    fig.update_layout(title="CET1 Ratio Under Stress Scenarios (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="CET1 Ratio (%)", xaxis_title="Quarter",
                      height=380)
    return fig

def plot_npa_evolution(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["gross_npa_ratio"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
            fill="tozeroy" if s_name == list(selected_scenarios)[0] else "none",
            fillcolor=hex_to_rgba(sc["color"], 0.08),
        ))
    fig.update_layout(title="Gross NPA Ratio Evolution (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="Gross NPA (%)", xaxis_title="Quarter", height=380)
    return fig

def plot_capital_waterfall(bank, scenario_name):
    sc = SCENARIOS[scenario_name]
    df, summary = run_stress_test(bank, sc)

    items = ["Starting CET1", "Credit Losses", "MTM / Market", "NII Impact",
             "Fee & Other", "Ending CET1"]
    credit_loss = -summary["cumulative_provisions"] / 100
    mtm_loss    = -summary["total_mtm_loss"] / 100
    nii_gain    = (df["nii"].sum() - bank["net_interest_income"] * 2) / 100
    fee_loss    = (df["fee_income"].sum() - bank["fee_income"] * 2) / 100
    start       = bank["cet1_capital"] / 100
    end         = summary["min_cet1"] * bank["rwa"] / 10000

    values = [start, credit_loss, mtm_loss, max(nii_gain, -1), max(fee_loss, -1), end]
    measures = ["absolute", "relative", "relative", "relative", "relative", "total"]
    colors = [COLORS["gold"], COLORS["red"], COLORS["orange"],
              COLORS["green"] if nii_gain >= 0 else COLORS["red"],
              COLORS["green"] if fee_loss >= 0 else COLORS["red"],
              COLORS["lightblue"]]

    fig = go.Figure(go.Waterfall(
        name="Capital Walk",
        orientation="v",
        measure=measures,
        x=items,
        y=values,
        connector=dict(line=dict(color=COLORS["midblue"], width=1.5)),
        increasing=dict(marker=dict(color=COLORS["green"])),
        decreasing=dict(marker=dict(color=COLORS["red"])),
        totals=dict(marker=dict(color=COLORS["gold"])),
        text=[f"₹{v*100:.0f}Cr" for v in values],
        textposition="outside",
        textfont=dict(color=COLORS["text"], size=11),
    ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        title=f"Capital Walk — {scenario_name.split(' ',1)[1]}",
        title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
        yaxis_title="₹ '00 Crore",
        height=380,
        margin=dict(l=40, r=20, t=50, b=40),
        showlegend=False,
    )
    fig.update_xaxes(tickfont=dict(color=COLORS["muted"]))
    fig.update_yaxes(gridcolor="#1e3a5f", tickfont=dict(color=COLORS["muted"]))
    return fig

def plot_sensitivity_heatmap(bank):
    """Create sensitivity heatmap: NPA multiplier vs. Rate shock."""
    npa_vals = [1.0, 1.3, 1.6, 2.0, 2.5, 3.0, 3.5]
    rate_vals = [0, 50, 100, 150, 200, 250, 300]
    base_sc = SCENARIOS["🔴 Severe Stress"].copy()
    z = []
    for nm in npa_vals:
        row = []
        for rb in rate_vals:
            sc = base_sc.copy()
            sc["npa_multiplier"] = nm
            sc["rate_shock_bps"] = rb
            _, s = run_stress_test(bank, sc)
            row.append(round(s["min_cet1"], 2))
        z.append(row)

    fig = go.Figure(go.Heatmap(
        z=z, x=[f"{r}bps" for r in rate_vals], y=[f"{n}x" for n in npa_vals],
        colorscale=[[0, "#dc3545"], [0.4, "#fd7e14"], [0.6, "#ffc107"],
                    [0.8, "#28a745"], [1.0, "#00cc66"]],
        zmin=6, zmax=12,
        text=[[f"{v:.1f}%" for v in row] for row in z],
        texttemplate="%{text}",
        textfont=dict(size=10, color="white"),
        colorbar=dict(title=dict(text="Min CET1 %", font=dict(color=COLORS["gold"])),
                      tickfont=dict(color=COLORS["text"]))
    ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        plot_bgcolor=PLOTLY_TEMPLATE["plot_bgcolor"],
        font=PLOTLY_TEMPLATE["font"],
        title="Sensitivity Heatmap: NPA Multiplier × Rate Shock → Min CET1 (%)",
        title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
        xaxis=dict(title="Rate Shock (bps)", tickfont=dict(color=COLORS["muted"]),
                   title_font=dict(color=COLORS["muted"])),
        yaxis=dict(title="NPA Multiplier", tickfont=dict(color=COLORS["muted"]),
                   title_font=dict(color=COLORS["muted"])),
        height=360,
        margin=dict(l=60, r=40, t=50, b=50),
    )
    return fig

def plot_liquidity_stress(bank, selected_scenarios):
    fig = base_fig()
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, _ = run_stress_test(bank, sc)
        fig.add_trace(go.Scatter(
            x=df["quarter"], y=df["lcr"],
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2.5),
            mode="lines+markers",
        ))
    fig.add_hline(y=100, line=dict(color="#dc3545", dash="dash", width=1.5),
                  annotation_text="LCR Min (100%)", annotation_font_color="#dc3545")
    fig.update_layout(title="Liquidity Coverage Ratio (LCR) Under Stress (%)",
                      title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                      yaxis_title="LCR (%)", xaxis_title="Quarter", height=380)
    return fig

def plot_pl_decomposition(bank, scenario_name):
    sc = SCENARIOS[scenario_name]
    df, _ = run_stress_test(bank, sc)
    fig = base_fig()
    fig.add_trace(go.Bar(x=df["quarter"], y=df["nii"],
                         name="NII", marker_color=COLORS["midblue"]))
    fig.add_trace(go.Bar(x=df["quarter"], y=df["fee_income"],
                         name="Fee Income", marker_color=COLORS["lightblue"]))
    fig.add_trace(go.Bar(x=df["quarter"], y=df["trading_income"],
                         name="Trading", marker_color=hex_to_rgba(COLORS["gold"], 0.53)))
    fig.add_trace(go.Bar(x=df["quarter"],
                         y=[-bank["operating_costs"]/4]*8,
                         name="Operating Costs", marker_color="#cc3333"))
    fig.add_trace(go.Bar(x=df["quarter"], y=-df["incremental_provisions"],
                         name="Provisions", marker_color="#ff6666"))
    fig.add_trace(go.Scatter(x=df["quarter"], y=df["net_profit"],
                             name="Net Profit", mode="lines+markers",
                             line=dict(color=COLORS["gold"], width=3),
                             marker=dict(size=8)))
    fig.update_layout(
        barmode="relative",
        title=f"P&L Decomposition — {scenario_name.split(' ',1)[1]}",
        title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
        yaxis_title="₹ Crore", xaxis_title="Quarter", height=400,
    )
    return fig

def plot_radar_scenarios(bank):
    cats = ["CET1 Ratio", "Peak NPA", "LCR", "NII Stability", "Net Profit"]
    fig = go.Figure()
    for s_name, sc in list(SCENARIOS.items())[1:5]:
        df, summary = run_stress_test(bank, sc)
        nii_stability = max(0, (df["nii"].min() / (bank["net_interest_income"]/4)) * 10)
        np_score = max(0, (df["net_profit"].sum() / bank["pat"] + 2) * 2.5)
        values = [
            min(summary["min_cet1"] - 5, 10),
            max(0, 15 - summary["peak_npa"]),
            min(summary["min_lcr"] / 20, 10),
            min(nii_stability, 10),
            min(np_score, 10),
        ]
        values += [values[0]]
        cats_closed = cats + [cats[0]]
        fig.add_trace(go.Scatterpolar(
            r=values, theta=cats_closed,
            name=s_name.split(" ", 1)[1],
            line=dict(color=sc["color"], width=2),
            fill="toself", fillcolor=hex_to_rgba(sc["color"], 0.13),
        ))
    fig.update_layout(
        paper_bgcolor=PLOTLY_TEMPLATE["paper_bgcolor"],
        polar=dict(
            bgcolor="#112240",
            radialaxis=dict(visible=True, range=[0, 10],
                            gridcolor="#1e3a5f", tickfont=dict(color=COLORS["muted"])),
            angularaxis=dict(gridcolor="#1e3a5f",
                             tickfont=dict(color=COLORS["text"]))
        ),
        title="Bank Resilience Radar (Higher = More Resilient)",
        title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
        legend=dict(bgcolor="#0a1628", bordercolor="#003366",
                    font=dict(color=COLORS["text"])),
        height=400,
        font=dict(color=COLORS["text"]),
        margin=dict(l=40, r=40, t=50, b=40),
    )
    return fig

# ─────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────
def metric_card(label, value, delta=None, delta_type="pos"):
    delta_html = ""
    if delta is not None:
        cls = f"delta-{delta_type}"
        arrow = "▲" if delta_type == "pos" else ("▼" if delta_type == "neg" else "●")
        delta_html = f'<div class="metric-delta {cls}">{arrow} {delta}</div>'
    return f"""
    <div class="metric-card">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        {delta_html}
    </div>"""

def risk_badge(level):
    levels = {"Low":"green","Moderate":"yellow","High":"orange","Critical":"red","Pass":"green","Breach":"red"}
    cls = levels.get(level, "yellow")
    return f'<span class="badge-{cls}">{level}</span>'

def section_header(title, icon=""):
    return f'<div class="section-header">{icon} {title}</div>'

def scenario_summary_table(bank, selected_scenarios):
    rows = []
    for s_name in selected_scenarios:
        sc = SCENARIOS[s_name]
        df, summary = run_stress_test(bank, sc)
        breach = "🔴 YES" if summary["any_breach"] else "🟢 No"
        rows.append({
            "Scenario": s_name,
            "Severity": sc["severity"],
            "Min CET1 (%)": f"{summary['min_cet1']:.2f}%",
            "Peak NPA (%)": f"{summary['peak_npa']:.1f}%",
            "Min LCR (%)": f"{summary['min_lcr']:.0f}%",
            "Cum. Provisions (₹Cr)": f"₹{summary['cumulative_provisions']:,.0f}",
            "Cum. Net Profit (₹Cr)": f"₹{summary['cumulative_net_profit']:,.0f}",
            "Regulatory Breach": breach,
        })
    df_summary = pd.DataFrame(rows)
    header_row = "".join(f"<th>{c}</th>" for c in df_summary.columns)
    data_rows = ""
    for _, row in df_summary.iterrows():
        data_rows += "<tr>" + "".join(f"<td>{v}</td>" for v in row) + "</tr>"
    return f"""
    <table class="styled-table">
        <thead><tr>{header_row}</tr></thead>
        <tbody>{data_rows}</tbody>
    </table>"""

# ─────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────
def main():
    apply_css()
    bank = generate_bank_data()

    # ── SIDEBAR ──────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"""
        <div style='text-align:center; padding:10px 0 5px 0;'>
            <div style='font-family:Playfair Display,serif;font-size:1.15rem;
                        font-weight:900;color:{COLORS["gold"]};letter-spacing:1px;'>
                🏦 THE MOUNTAIN PATH
            </div>
            <div style='color:{COLORS["lightblue"]};font-size:0.72rem;
                        letter-spacing:2px;text-transform:uppercase;margin-top:2px;'>
                Bank Stress Testing Lab
            </div>
            <hr style='border-color:{COLORS["gold"]}44;margin:10px 0;'>
        </div>
        """, unsafe_allow_html=True)

        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>📋 NAVIGATION</div>", unsafe_allow_html=True)
        page = st.selectbox("", [
            "🏠 Dashboard",
            "📊 Balance Sheet & Income",
            "🎯 Run Stress Tests",
            "📈 Capital Analysis",
            "💧 Liquidity Analysis",
            "🔥 Sensitivity & Heatmap",
            "🛠️ Custom Scenario Builder",
            "📋 Scenario Comparison",
            "🔄 Reverse Stress Test",
        ], label_visibility="collapsed")

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>📌 SELECT SCENARIOS</div>", unsafe_allow_html=True)
        selected_scenarios = []
        for s_name, sc in SCENARIOS.items():
            checked = st.checkbox(s_name, value=(s_name in ["🟢 Baseline", "🟠 Moderate Stress", "🔴 Severe Stress"]),
                                  key=f"cb_{s_name}")
            if checked:
                selected_scenarios.append(s_name)
        if not selected_scenarios:
            selected_scenarios = ["🟢 Baseline"]

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;font-size:0.85rem;'>⚙️ BASELINE PARAMETERS</div>", unsafe_allow_html=True)
        bank["gross_npa_ratio"] = st.slider("Starting Gross NPA (%)", 3.0, 18.0, 7.0, 0.5)
        bank["cet1_ratio"]      = st.slider("Starting CET1 Ratio (%)", 8.0, 16.0, 9.57, 0.1)
        bank["lcr"]             = st.slider("Starting LCR (%)", 100.0, 200.0, 142.0, 5.0)
        bank["nim"]             = st.slider("Starting NIM (%)", 2.0, 5.5, 3.9, 0.1)

        st.markdown("<hr style='border-color:#1e3a5f;'>", unsafe_allow_html=True)
        st.markdown(f"""
        <div style='color:{COLORS['muted']};font-size:0.72rem;text-align:center;'>
            Prof. V. Ravichandran<br>
            <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
               style='color:{COLORS["gold"]};'>LinkedIn</a> &nbsp;|&nbsp;
            <a href='https://github.com/trichyravis' target='_blank'
               style='color:{COLORS["gold"]};'>GitHub</a>
        </div>""", unsafe_allow_html=True)

    # ── HEADER ───────────────────────────────────────────────────
    st.markdown(f"""
    <div style='padding:10px 0 18px 0;'>
        <div class='brand-tag'>The Mountain Path — World of Finance</div>
        <div class='hero-title'>🏦 Bank Stress Testing Lab</div>
        <div class='hero-subtitle'>{bank["name"]} &nbsp;|&nbsp; {bank["type"]}
            &nbsp;|&nbsp; Rating: <span style='color:{COLORS["gold"]}'>{bank["rating"]}</span>
        </div>
    </div>
    <hr class='gold-divider'>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 1: DASHBOARD
    # ═══════════════════════════════════════════════════════════════
    if page == "🏠 Dashboard":
        st.markdown(section_header("Key Risk Indicators — Current Position", "📊"), unsafe_allow_html=True)

        cols = st.columns(5)
        metrics = [
            ("CET1 Ratio", f"{bank['cet1_ratio']:.1f}%", "Above 8% min", "pos"),
            ("Gross NPA", f"{bank['gross_npa_ratio']:.1f}%", "Moderate risk", "warn"),
            ("LCR", f"{bank['lcr']:.0f}%", "Above 100% min", "pos"),
            ("CRAR", f"{bank['crar']:.1f}%", "Above 11.5% min", "pos"),
            ("NIM", f"{bank['nim']:.1f}%", "Healthy spread", "pos"),
        ]
        for col, (lbl, val, delta, dtype) in zip(cols, metrics):
            with col:
                st.markdown(metric_card(lbl, val, delta, dtype), unsafe_allow_html=True)

        cols2 = st.columns(5)
        metrics2 = [
            ("Total Assets", f"₹{bank['total_assets']:,}Cr", None, "pos"),
            ("Gross Loans", f"₹{bank['gross_loans']:,}Cr", "72.9% C/D ratio", "pos"),
            ("CASA Ratio", f"{bank['casa_ratio']:.0f}%", "Good funding mix", "pos"),
            ("Net Interest Income", f"₹{bank['net_interest_income']:,}Cr", "Annual", "pos"),
            ("Return on Equity", f"{bank['roe']:.1f}%", "FY Annualised", "pos"),
        ]
        for col, (lbl, val, delta, dtype) in zip(cols2, metrics2):
            with col:
                st.markdown(metric_card(lbl, val, delta, dtype), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(section_header("CET1 Under All Scenarios", "💰"), unsafe_allow_html=True)
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios),
                            use_container_width=True)
        with c2:
            st.markdown(section_header("NPA Evolution", "📉"), unsafe_allow_html=True)
            st.plotly_chart(plot_npa_evolution(bank, selected_scenarios),
                            use_container_width=True)

        st.markdown(section_header("Quick Scenario Summary", "🎯"), unsafe_allow_html=True)
        st.markdown(scenario_summary_table(bank, selected_scenarios), unsafe_allow_html=True)

        # Radar
        st.markdown("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns([1.2, 0.8])
        with c1:
            st.markdown(section_header("Resilience Radar", "🕸️"), unsafe_allow_html=True)
            st.plotly_chart(plot_radar_scenarios(bank), use_container_width=True)
        with c2:
            st.markdown(section_header("Sector Concentration Risk", "🏗️"), unsafe_allow_html=True)
            sectors = {
                "Real Estate": bank["sector_real_estate"],
                "Infrastructure": bank["sector_infrastructure"],
                "NBFC": bank["sector_nbfc"],
                "Power": bank["sector_power"],
                "Textile": bank["sector_textile"],
                "Gems/Jewellery": bank["sector_gems"],
            }
            fig_pie = go.Figure(go.Pie(
                labels=list(sectors.keys()),
                values=list(sectors.values()),
                hole=0.5,
                marker=dict(colors=[COLORS["darkblue"], COLORS["midblue"],
                                    COLORS["gold"], COLORS["lightblue"],
                                    COLORS["orange"], COLORS["green"]]),
                textfont=dict(color="white", size=11),
            ))
            fig_pie.update_layout(
                paper_bgcolor="#0a1628",
                font=dict(color=COLORS["text"]),
                legend=dict(bgcolor="#0a1628", font=dict(color=COLORS["text"])),
                margin=dict(l=10, r=10, t=30, b=10),
                height=380,
            )
            st.plotly_chart(fig_pie, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 2: BALANCE SHEET & INCOME
    # ═══════════════════════════════════════════════════════════════
    elif page == "📊 Balance Sheet & Income":
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(section_header("Balance Sheet — Assets (₹ Crore)", "📋"), unsafe_allow_html=True)
            bs_assets = pd.DataFrame({
                "Item": ["Gross Loans & Advances", "  ► Retail", "  ► Corporate",
                          "  ► MSME", "  ► Agriculture",
                          "HTM Investments", "AFS Investments", "Equity Portfolio",
                          "Cash & HQLA", "Fixed Assets", "Other Assets"],
                "Amount (₹Cr)": [
                    bank["gross_loans"], bank["retail_loans"], bank["corporate_loans"],
                    bank["msme_loans"], bank["agri_loans"],
                    bank["investments_htm"], bank["investments_afs"],
                    bank["equity_portfolio"], bank["cash_hqla"],
                    bank["fixed_assets"], bank["other_assets"]
                ],
                "% of Assets": [
                    f"{v/bank['total_assets']*100:.1f}%" for v in [
                        bank["gross_loans"], bank["retail_loans"], bank["corporate_loans"],
                        bank["msme_loans"], bank["agri_loans"],
                        bank["investments_htm"], bank["investments_afs"],
                        bank["equity_portfolio"], bank["cash_hqla"],
                        bank["fixed_assets"], bank["other_assets"]
                    ]
                ]
            })
            st.dataframe(bs_assets, use_container_width=True, height=380, hide_index=True)

        with c2:
            st.markdown(section_header("Balance Sheet — Liabilities & Capital (₹ Crore)", "📋"), unsafe_allow_html=True)
            bs_liab = pd.DataFrame({
                "Item": ["Total Deposits", "  ► CASA Deposits", "  ► Term Deposits",
                          "Wholesale Funding", "Subordinated Debt",
                          "CET1 Capital", "Additional Tier 1", "Tier 2 Capital",
                          "Total Capital"],
                "Amount (₹Cr)": [
                    bank["total_deposits"], bank["casa_deposits"], bank["term_deposits"],
                    bank["wholesale_funding"], bank["sub_debt"],
                    bank["cet1_capital"], bank["tier1_capital"] - bank["cet1_capital"],
                    bank["tier2_capital"], bank["total_capital"]
                ],
            })
            st.dataframe(bs_liab, use_container_width=True, height=380, hide_index=True)

        c3, c4 = st.columns(2)
        with c3:
            st.markdown(section_header("Income Statement — Annual (₹ Crore)", "💼"), unsafe_allow_html=True)
            is_data = pd.DataFrame({
                "Line Item": ["Net Interest Income", "Fee & Commission Income",
                               "Trading & MTM Income", "Other Income",
                               "Total Operating Income",
                               "Operating Costs", "Pre-Provision Operating Profit",
                               "Provisions & Write-offs", "Profit Before Tax",
                               "Taxes", "Profit After Tax"],
                "Amount (₹Cr)": [
                    bank["net_interest_income"], bank["fee_income"],
                    bank["trading_income"], bank["other_income"],
                    bank["net_interest_income"] + bank["fee_income"] + bank["trading_income"] + bank["other_income"],
                    -bank["operating_costs"], bank["pre_provision_profit"],
                    -bank["provisions_charge"], bank["pbt"], -bank["tax"], bank["pat"]
                ],
            })
            st.dataframe(is_data, use_container_width=True, height=380, hide_index=True)

        with c4:
            st.markdown(section_header("Capital & Asset Quality Ratios", "📐"), unsafe_allow_html=True)
            ratios = pd.DataFrame({
                "Ratio": ["CET1 Ratio", "Tier 1 Ratio", "CRAR (Total)", "Leverage Ratio",
                           "Gross NPA Ratio", "Net NPA Ratio", "SMA Ratio",
                           "Restructured Assets", "PCR",
                           "NIM", "ROE", "ROA", "CASA Ratio", "C/D Ratio",
                           "LCR", "NSFR"],
                "Value": ["9.57%","10.64%","12.90%","8.10%",
                          "7.00%","3.50%","3.20%","1.80%","50%",
                          "3.90%","13.70%","1.10%","40.0%","72.9%",
                          "142%","118%"],
                "Regulatory Min": ["8.0%","9.5%","11.5%","—",
                                    "—","—","—","—","—",
                                    "—","—","—","—","—",
                                    "100%","100%"],
                "Status": ["✅ Pass","✅ Pass","✅ Pass","✅ Pass",
                            "⚠️ Watch","⚠️ Watch","⚠️ Watch","⚠️ Watch","✅ Pass",
                            "✅ Pass","✅ Pass","✅ Pass","✅ Pass","✅ Pass",
                            "✅ Pass","✅ Pass"],
            })
            st.dataframe(ratios, use_container_width=True, height=450, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 3: RUN STRESS TESTS
    # ═══════════════════════════════════════════════════════════════
    elif page == "🎯 Run Stress Tests":
        sc_choice = st.selectbox("Select Scenario to Analyse in Detail", list(SCENARIOS.keys()))
        sc = SCENARIOS[sc_choice]
        df, summary = run_stress_test(bank, sc)

        st.markdown(f"""
        <div class='metric-card' style='margin-bottom:16px;'>
            <b style='color:{COLORS["gold"]};font-size:1.1rem;'>{sc_choice}</b><br>
            <span style='color:{COLORS["muted"]};'>{sc["description"]}</span>&nbsp;&nbsp;
            <span class='scenario-pill'>Severity: {sc["severity"]}</span>
        </div>""", unsafe_allow_html=True)

        # Shock parameters
        st.markdown(section_header("Scenario Shock Parameters", "⚡"), unsafe_allow_html=True)
        cols = st.columns(4)
        params = [
            ("Rate Shock", f"{sc['rate_shock_bps']} bps"),
            ("Equity Shock", f"{sc['equity_shock_pct']*100:.0f}%"),
            ("NPA Multiplier", f"{sc['npa_multiplier']}×"),
            ("INR Depreciation", f"{sc['inr_depreciation']*100:.0f}%"),
            ("Property Shock", f"{sc['property_shock']*100:.0f}%"),
            ("Deposit Run-off", f"{sc['deposit_runoff']*100:.0f}%"),
            ("Fee Income Shock", f"{sc['fee_income_shock']*100:.0f}%"),
            ("Credit Spread Widening", f"{sc['credit_spread_widening']} bps"),
        ]
        for i, (lbl, val) in enumerate(params):
            with cols[i % 4]:
                st.markdown(metric_card(lbl, val), unsafe_allow_html=True)

        # Results summary
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Stress Test Results", "📊"), unsafe_allow_html=True)
        r_cols = st.columns(4)
        breach_status = "🔴 BREACH" if summary["any_breach"] else "🟢 PASS"
        with r_cols[0]:
            st.markdown(metric_card("Min CET1 Ratio", f"{summary['min_cet1']:.2f}%",
                "Below 8% min!" if summary["min_cet1"] < 8 else "Above regulatory min",
                "neg" if summary["min_cet1"] < 8 else "pos"), unsafe_allow_html=True)
        with r_cols[1]:
            st.markdown(metric_card("Peak Gross NPA", f"{summary['peak_npa']:.1f}%",
                "Stress peak", "neg"), unsafe_allow_html=True)
        with r_cols[2]:
            st.markdown(metric_card("Min LCR", f"{summary['min_lcr']:.0f}%",
                "Below 100!" if summary["min_lcr"] < 100 else "Above min",
                "neg" if summary["min_lcr"] < 100 else "pos"), unsafe_allow_html=True)
        with r_cols[3]:
            st.markdown(metric_card("Regulatory Status", breach_status,
                f"{summary['cet1_breaches']} CET1 breach qtr(s)",
                "neg" if summary["any_breach"] else "pos"), unsafe_allow_html=True)

        # Charts row
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_pl_decomposition(bank, sc_choice), use_container_width=True)
        with c2:
            st.plotly_chart(plot_capital_waterfall(bank, sc_choice), use_container_width=True)

        # Quarterly data table
        st.markdown(section_header("Quarterly Projection Detail", "📅"), unsafe_allow_html=True)
        display_df = df[["quarter","gdp_growth","gross_npa_ratio","incremental_provisions",
                          "nii","net_profit","cet1_ratio","crar","lcr"]].copy()
        display_df.columns = ["Quarter","GDP Growth (%)","Gross NPA (%)","Provisions (₹Cr)",
                                "NII (₹Cr)","Net Profit (₹Cr)","CET1 (%)","CRAR (%)","LCR (%)"]
        display_df = display_df.round(2)
        st.dataframe(display_df, use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 4: CAPITAL ANALYSIS
    # ═══════════════════════════════════════════════════════════════
    elif page == "📈 Capital Analysis":
        st.markdown(section_header("Capital Ratio Paths Under All Selected Scenarios", "💰"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios), use_container_width=True)
        with c2:
            # CRAR
            fig = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig.add_trace(go.Scatter(x=df["quarter"], y=df["crar"],
                    name=s_name.split(" ",1)[1],
                    line=dict(color=sc["color"], width=2.5), mode="lines+markers"))
            fig.add_hline(y=11.5, line=dict(color="#dc3545", dash="dash", width=1.5),
                          annotation_text="CRAR Min (11.5%)")
            fig.update_layout(title="Total CRAR Under Stress (%)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                yaxis_title="CRAR (%)", height=380)
            st.plotly_chart(fig, use_container_width=True)

        # Waterfall for all selected
        st.markdown(section_header("Capital Walk — Waterfall Charts", "🏗️"), unsafe_allow_html=True)
        wf_cols = st.columns(min(len(selected_scenarios), 3))
        for i, s_name in enumerate(selected_scenarios[:3]):
            with wf_cols[i % 3]:
                st.plotly_chart(plot_capital_waterfall(bank, s_name), use_container_width=True)

        # Capital buffer analysis
        st.markdown(section_header("Capital Buffer Analysis", "🛡️"), unsafe_allow_html=True)
        buf_data = []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            _, summary = run_stress_test(bank, sc)
            buf_data.append({
                "Scenario": s_name,
                "Min CET1 (%)": round(summary["min_cet1"], 2),
                "Buffer above 8% floor (%)": round(summary["min_cet1"] - 8.0, 2),
                "Buffer above 9.5% incl. conservation (%)": round(summary["min_cet1"] - 9.5, 2),
                "Capital at Risk (₹Cr)": round(
                    max(0, (bank["cet1_ratio"] - summary["min_cet1"]) / 100 * bank["rwa"]), 0),
                "Pass / Fail CET1": "✅ Pass" if summary["min_cet1"] >= 8.0 else "❌ Fail",
            })
        st.dataframe(pd.DataFrame(buf_data), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 5: LIQUIDITY ANALYSIS
    # ═══════════════════════════════════════════════════════════════
    elif page == "💧 Liquidity Analysis":
        st.markdown(section_header("Liquidity Stress Testing", "💧"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_liquidity_stress(bank, selected_scenarios), use_container_width=True)
        with c2:
            # HQLA evolution
            fig = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig.add_trace(go.Bar(x=df["quarter"], y=df["hqla"],
                    name=s_name.split(" ",1)[1], marker_color=sc["color"]))
            fig.update_layout(title="HQLA Buffer (₹ Crore)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                yaxis_title="₹ Crore", barmode="group", height=380)
            st.plotly_chart(fig, use_container_width=True)

        st.markdown(section_header("Deposit Run-off Analysis", "🏃"), unsafe_allow_html=True)
        fig_dep = base_fig()
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            df, _ = run_stress_test(bank, sc)
            fig_dep.add_trace(go.Scatter(x=df["quarter"], y=df["deposits"],
                name=s_name.split(" ",1)[1],
                line=dict(color=sc["color"], width=2.5), mode="lines+markers"))
        fig_dep.update_layout(title="Total Deposits Under Stress (₹ Crore)",
            title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
            yaxis_title="₹ Crore", height=360)
        st.plotly_chart(fig_dep, use_container_width=True)

        # Liquidity metrics table
        liq_data = []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            df, summary = run_stress_test(bank, sc)
            liq_data.append({
                "Scenario": s_name,
                "Starting LCR (%)": bank["lcr"],
                "Min LCR (%)": round(summary["min_lcr"], 0),
                "Deposit Run-off (%)": f"{sc['deposit_runoff']*100:.0f}%",
                "HQLA Floor (₹Cr)": round(df["hqla"].min(), 0),
                "Liquidity Status": "❌ Breach" if summary["min_lcr"] < 100 else "✅ Pass",
            })
        st.dataframe(pd.DataFrame(liq_data), use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 6: SENSITIVITY & HEATMAP
    # ═══════════════════════════════════════════════════════════════
    elif page == "🔥 Sensitivity & Heatmap":
        st.markdown(section_header("Sensitivity Heatmap: NPA Multiplier × Rate Shock", "🔥"), unsafe_allow_html=True)
        st.plotly_chart(plot_sensitivity_heatmap(bank), use_container_width=True)

        st.markdown(section_header("Single-Variable Sensitivity Analysis", "📊"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        base_sc = SCENARIOS["🔴 Severe Stress"].copy()
        with c1:
            npa_vals = np.linspace(1.0, 4.0, 15)
            npa_sens = run_sensitivity(bank, "npa_multiplier", npa_vals, base_sc)
            fig = base_fig()
            fig.add_trace(go.Scatter(x=npa_vals, y=npa_sens["min_cet1"],
                name="Min CET1", line=dict(color=COLORS["gold"], width=2.5),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["gold"], 0.08)))
            fig.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                          annotation_text="CET1 Min 8%")
            fig.update_layout(title="CET1 Sensitivity to NPA Multiplier",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                xaxis_title="NPA Multiplier", yaxis_title="Min CET1 (%)", height=360)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            rate_vals = np.linspace(0, 400, 15)
            rate_sens = run_sensitivity(bank, "rate_shock_bps", rate_vals, base_sc)
            fig2 = base_fig()
            fig2.add_trace(go.Scatter(x=rate_vals, y=rate_sens["min_cet1"],
                name="Min CET1", line=dict(color=COLORS["lightblue"], width=2.5),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["lightblue"], 0.08)))
            fig2.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                           annotation_text="CET1 Min 8%")
            fig2.update_layout(title="CET1 Sensitivity to Rate Shock (bps)",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                xaxis_title="Rate Shock (bps)", yaxis_title="Min CET1 (%)", height=360)
            st.plotly_chart(fig2, use_container_width=True)

        # Equity shock sensitivity
        eq_vals = np.linspace(0, 0.70, 15)
        eq_sens = run_sensitivity(bank, "equity_shock_pct", eq_vals, base_sc)
        fig3 = base_fig()
        fig3.add_trace(go.Scatter(x=eq_vals*100, y=eq_sens["min_cet1"],
            name="Min CET1", line=dict(color=COLORS["orange"], width=2.5),
            fill="tozeroy", fillcolor=hex_to_rgba(COLORS["orange"], 0.08)))
        fig3.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"),
                       annotation_text="CET1 Min 8%")
        fig3.update_layout(title="CET1 Sensitivity to Equity Market Crash (%)",
            title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
            xaxis_title="Equity Shock (%)", yaxis_title="Min CET1 (%)", height=350)
        st.plotly_chart(fig3, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 7: CUSTOM SCENARIO BUILDER
    # ═══════════════════════════════════════════════════════════════
    elif page == "🛠️ Custom Scenario Builder":
        st.markdown(section_header("Build Your Own Stress Scenario", "🛠️"), unsafe_allow_html=True)

        with st.form("custom_scenario"):
            col1, col2, col3 = st.columns(3)
            with col1:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>📉 Credit Risk</div>", unsafe_allow_html=True)
                npa_mult = st.slider("NPA Multiplier", 1.0, 5.0, 2.0, 0.1,
                                     help="How many times NPAs grow vs current")
                deposit_ro = st.slider("Deposit Run-off (%)", 0.0, 30.0, 8.0, 0.5) / 100
                property_sh = st.slider("Property Price Shock (%)", 0.0, 50.0, 20.0, 1.0) / 100

            with col2:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>📈 Market Risk</div>", unsafe_allow_html=True)
                rate_sh = st.slider("Rate Shock (bps)", 0, 500, 200, 25)
                eq_sh = st.slider("Equity Market Fall (%)", 0.0, 70.0, 30.0, 1.0) / 100
                inr_sh = st.slider("INR Depreciation (%)", 0.0, 40.0, 10.0, 1.0) / 100

            with col3:
                st.markdown(f"<div style='color:{COLORS['gold']};font-weight:700;'>🌍 Macro Variables</div>", unsafe_allow_html=True)
                gdp_shock_val = st.slider("GDP Shock (pp, peak quarter)", 0.0, 10.0, 3.5, 0.1)
                fee_sh = st.slider("Fee Income Fall (%)", 0.0, 50.0, 20.0, 1.0) / 100
                cs_wide = st.slider("Credit Spread Widening (bps)", 0, 600, 200, 25)
                scenario_name_custom = st.text_input("Scenario Name", "My Custom Scenario")

            submitted = st.form_submit_button("🚀 Run Custom Stress Test",
                                              use_container_width=True)

        if submitted:
            gdp_path = np.array([0, -gdp_shock_val*0.5, -gdp_shock_val, -gdp_shock_val*0.9,
                                  -gdp_shock_val*0.7, -gdp_shock_val*0.4,
                                  -gdp_shock_val*0.2, -gdp_shock_val*0.1])
            custom_sc = {
                "color": "#cc99ff",
                "severity": "Custom",
                "description": scenario_name_custom,
                "gdp_shock": gdp_path.tolist(),
                "rate_shock_bps": rate_sh,
                "equity_shock_pct": -eq_sh,
                "inr_depreciation": inr_sh,
                "npa_multiplier": npa_mult,
                "deposit_runoff": deposit_ro,
                "property_shock": -property_sh,
                "credit_spread_widening": cs_wide,
                "fee_income_shock": -fee_sh,
            }
            df_c, summary_c = run_stress_test(bank, custom_sc)

            st.success(f"✅ Custom scenario '{scenario_name_custom}' executed successfully!")

            r1, r2, r3, r4 = st.columns(4)
            with r1: st.markdown(metric_card("Min CET1", f"{summary_c['min_cet1']:.2f}%",
                "BREACH!" if summary_c['min_cet1'] < 8 else "Pass",
                "neg" if summary_c['min_cet1'] < 8 else "pos"), unsafe_allow_html=True)
            with r2: st.markdown(metric_card("Peak NPA", f"{summary_c['peak_npa']:.1f}%"), unsafe_allow_html=True)
            with r3: st.markdown(metric_card("Min LCR", f"{summary_c['min_lcr']:.0f}%",
                "BREACH!" if summary_c['min_lcr'] < 100 else "Pass",
                "neg" if summary_c['min_lcr'] < 100 else "pos"), unsafe_allow_html=True)
            with r4: st.markdown(metric_card("Cum. Provisions",
                f"₹{summary_c['cumulative_provisions']:,.0f}Cr"), unsafe_allow_html=True)

            c_plot1, c_plot2 = st.columns(2)
            with c_plot1:
                st.plotly_chart(plot_pl_decomposition(bank, "🟢 Baseline"), use_container_width=True)
            with c_plot2:
                # Custom capital path
                fig_cap = base_fig()
                fig_cap.add_trace(go.Scatter(x=df_c["quarter"], y=df_c["cet1_ratio"],
                    name=scenario_name_custom, line=dict(color="#cc99ff", width=3),
                    mode="lines+markers", marker=dict(size=8)))
                bl_df, _ = run_stress_test(bank, SCENARIOS["🟢 Baseline"])
                fig_cap.add_trace(go.Scatter(x=bl_df["quarter"], y=bl_df["cet1_ratio"],
                    name="Baseline", line=dict(color=COLORS["green"], dash="dot", width=2)))
                fig_cap.add_hline(y=8.0, line=dict(color=COLORS["red"], dash="dash"))
                fig_cap.update_layout(title="CET1 Ratio: Custom vs Baseline",
                    title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                    height=380)
                st.plotly_chart(fig_cap, use_container_width=True)

            st.dataframe(df_c[["quarter","gdp_growth","gross_npa_ratio",
                                "net_profit","cet1_ratio","lcr"]].round(2),
                         use_container_width=True, hide_index=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 8: SCENARIO COMPARISON
    # ═══════════════════════════════════════════════════════════════
    elif page == "📋 Scenario Comparison":
        st.markdown(section_header("Full Scenario Comparison Table", "📋"), unsafe_allow_html=True)
        st.markdown(scenario_summary_table(bank, selected_scenarios), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Side-by-Side Capital & NPA Paths", "📊"), unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            st.plotly_chart(plot_multi_scenario_cet1(bank, selected_scenarios), use_container_width=True)
        with c2:
            st.plotly_chart(plot_npa_evolution(bank, selected_scenarios), use_container_width=True)

        c3, c4 = st.columns(2)
        with c3:
            st.plotly_chart(plot_liquidity_stress(bank, selected_scenarios), use_container_width=True)
        with c4:
            # Net profit comparison
            fig_np = base_fig()
            for s_name in selected_scenarios:
                sc = SCENARIOS[s_name]
                df, _ = run_stress_test(bank, sc)
                fig_np.add_trace(go.Bar(x=df["quarter"], y=df["net_profit"],
                    name=s_name.split(" ",1)[1], marker_color=sc["color"]))
            fig_np.update_layout(title="Quarterly Net Profit Under Stress (₹Cr)",
                title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
                barmode="group", yaxis_title="₹ Crore", height=380)
            st.plotly_chart(fig_np, use_container_width=True)

        # Provisions comparison
        st.markdown(section_header("Cumulative Provisions by Scenario (₹ Crore)", "💸"), unsafe_allow_html=True)
        fig_prov = base_fig()
        labels, prov_vals, colors_list = [], [], []
        for s_name in selected_scenarios:
            sc = SCENARIOS[s_name]
            _, summary = run_stress_test(bank, sc)
            labels.append(s_name.split(" ",1)[1])
            prov_vals.append(summary["cumulative_provisions"])
            colors_list.append(sc["color"])
        fig_prov.add_trace(go.Bar(x=labels, y=prov_vals,
            marker_color=colors_list, text=[f"₹{v:,.0f}Cr" for v in prov_vals],
            textposition="outside", textfont=dict(color=COLORS["text"])))
        fig_prov.update_layout(title="Cumulative Provisions (8Q)",
            title_font=dict(color=COLORS["gold"], size=14, family="Playfair Display"),
            yaxis_title="₹ Crore", height=380)
        st.plotly_chart(fig_prov, use_container_width=True)

    # ═══════════════════════════════════════════════════════════════
    # PAGE 9: REVERSE STRESS TEST
    # ═══════════════════════════════════════════════════════════════
    elif page == "🔄 Reverse Stress Test":
        st.markdown(section_header("Reverse Stress Testing — What Breaks the Bank?", "🔄"), unsafe_allow_html=True)
        st.markdown(f"""
        <div class='metric-card'>
            <b style='color:{COLORS["gold"]}'>What is Reverse Stress Testing?</b><br>
            <span style='color:{COLORS["muted"]};'>Instead of asking "what happens if GDP falls 5%?",
            reverse stress testing asks: <i>"what combination of shocks reduces CET1 below 8%?"</i>
            It identifies the exact Achilles' heel of the balance sheet.</span>
        </div>""", unsafe_allow_html=True)

        st.markdown(section_header("Critical Threshold Analysis", "⚠️"), unsafe_allow_html=True)

        # Find the NPA multiplier that breaks CET1
        base_sc = SCENARIOS["🟠 Moderate Stress"].copy()
        npa_breaks = None
        for nm in np.linspace(1.0, 5.0, 100):
            sc_test = base_sc.copy()
            sc_test["npa_multiplier"] = nm
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                npa_breaks = nm
                break

        rate_breaks = None
        for rb in range(0, 600, 5):
            sc_test = base_sc.copy()
            sc_test["rate_shock_bps"] = rb
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                rate_breaks = rb
                break

        equity_breaks = None
        for eq in np.linspace(0, 1.0, 100):
            sc_test = base_sc.copy()
            sc_test["equity_shock_pct"] = -eq
            _, s = run_stress_test(bank, sc_test)
            if s["min_cet1"] < 8.0:
                equity_breaks = eq
                break

        c1, c2, c3 = st.columns(3)
        with c1:
            val = f"{npa_breaks:.1f}×" if npa_breaks else "Resilient >5×"
            st.markdown(metric_card("NPA Break-Even Multiplier",
                val, "CET1 breaches at this level",
                "neg" if npa_breaks and npa_breaks < 3 else "warn"), unsafe_allow_html=True)
        with c2:
            val2 = f"{rate_breaks}bps" if rate_breaks else "Resilient >600bps"
            st.markdown(metric_card("Rate Shock Break-Even",
                val2, "CET1 breaches at this level",
                "neg" if rate_breaks and rate_breaks < 200 else "warn"), unsafe_allow_html=True)
        with c3:
            val3 = f"{equity_breaks*100:.0f}%" if equity_breaks else "Resilient >100%"
            st.markdown(metric_card("Equity Crash Break-Even",
                val3, "CET1 breaches at this level",
                "neg" if equity_breaks and equity_breaks < 0.4 else "warn"), unsafe_allow_html=True)

        # Break-even visualisation
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(section_header("Break-Even Frontier: NPA × Rate Shock", "📍"), unsafe_allow_html=True)

        npa_range = np.linspace(1.0, 4.5, 20)
        rate_range = np.linspace(0, 400, 20)
        frontier_npa, frontier_rate = [], []
        for nm in npa_range:
            for rb in rate_range:
                sc_test = base_sc.copy()
                sc_test["npa_multiplier"] = nm
                sc_test["rate_shock_bps"] = int(rb)
                _, s = run_stress_test(bank, sc_test)
                if s["min_cet1"] < 8.0:
                    frontier_npa.append(nm)
                    frontier_rate.append(rb)
                    break

        if frontier_npa:
            fig_rev = base_fig()
            fig_rev.add_trace(go.Scatter(
                x=frontier_npa, y=frontier_rate,
                mode="lines+markers",
                line=dict(color=COLORS["red"], width=3),
                fill="tozeroy", fillcolor=hex_to_rgba(COLORS["red"], 0.13),
                name="Failure Frontier",
                marker=dict(size=6, color=COLORS["red"]),
            ))
            fig_rev.add_annotation(
                x=np.mean(frontier_npa), y=np.mean(frontier_rate) * 0.4,
                text="BANK SURVIVES ✅",
                font=dict(color=COLORS["green"], size=14, family="Playfair Display"),
                showarrow=False,
            )
            fig_rev.add_annotation(
                x=np.mean(frontier_npa), y=np.mean(frontier_rate) * 1.5,
                text="CET1 BREACH ❌",
                font=dict(color=COLORS["red"], size=14, family="Playfair Display"),
                showarrow=False,
            )
            fig_rev.update_layout(
                title="Reverse Stress: Failure Frontier (NPA Multiplier vs Rate Shock)",
                title_font=dict(color=COLORS["gold"], size=15, family="Playfair Display"),
                xaxis_title="NPA Multiplier",
                yaxis_title="Rate Shock (bps)",
                height=420,
            )
            st.plotly_chart(fig_rev, use_container_width=True)

        # Management actions
        st.markdown(section_header("Management Actions Available", "🔧"), unsafe_allow_html=True)
        actions = pd.DataFrame({
            "Action": [
                "Capital raising via QIP/Rights issue",
                "Accelerate NPA recoveries / NARCL referral",
                "Reduce risk appetite — cut new corporate lending",
                "Sell AFS portfolio to crystallise gains early",
                "Increase PCR to 70% pre-emptively",
                "Activate contingent liquidity facilities (RBI SLF)",
                "Increase CASA mobilisation",
                "Reduce dividend payout ratio",
            ],
            "Capital Impact (₹Cr est.)": ["+3,000–5,000", "+500–1,200", "+800 RWA relief",
                                          "+400–700", "-300 provisions", "Liquidity only",
                                          "Funding cost -15bps", "+200–400"],
            "Timeline": ["3–6 months","6–12 months","Immediate","1–2 months",
                          "Immediate","Immediate","6–18 months","Annual"],
            "Priority": ["⭐⭐⭐ Critical","⭐⭐⭐ Critical","⭐⭐ High","⭐⭐ High",
                          "⭐⭐ High","⭐ Medium","⭐ Medium","⭐ Medium"],
        })
        st.dataframe(actions, use_container_width=True, hide_index=True)

    # ── FOOTER ───────────────────────────────────────────────────
    st.markdown("<hr class='gold-divider'>", unsafe_allow_html=True)
    st.markdown(f"""
    <div style='text-align:center; padding:8px 0; color:{COLORS["muted"]}; font-size:0.78rem;'>
        <span style='color:{COLORS["gold"]};font-weight:700;font-family:Playfair Display,serif;'>
            THE MOUNTAIN PATH — World of Finance
        </span><br>
        Bank Stress Testing Lab &nbsp;|&nbsp; Prof. V. Ravichandran &nbsp;|&nbsp;
        28+ Yrs Corporate Finance &amp; Banking &nbsp;|&nbsp;
        <a href='https://www.linkedin.com/in/trichyravis' target='_blank'
           style='color:{COLORS["gold"]};text-decoration:none;'>LinkedIn</a>
        &nbsp;|&nbsp;
        <a href='https://github.com/trichyravis' target='_blank'
           style='color:{COLORS["gold"]};text-decoration:none;'>GitHub</a>
    </div>""", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
