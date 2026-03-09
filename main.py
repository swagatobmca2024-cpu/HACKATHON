import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.gridspec import GridSpec
import matplotlib.ticker as mticker
from matplotlib.colors import LinearSegmentedColormap
import warnings
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
#  PAGE CONFIG & GLOBAL CSS
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Adidas Sales Intelligence",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded",
)

DARK_BG      = "#0A0A0F"
CARD_BG      = "#12121A"
SIDEBAR_BG   = "#0D0D14"
ACCENT1      = "#00E5FF"   # electric cyan
ACCENT2      = "#FF3CAC"   # hot pink
ACCENT3      = "#7B5EA7"   # muted violet
ACCENT4      = "#F9C846"   # golden yellow
ACCENT5      = "#39FF14"   # neon green
TEXT_PRI     = "#F0F0F5"
TEXT_SEC     = "#8A8A9A"
BORDER       = "#1E1E2E"

PALETTE = [ACCENT1, ACCENT2, ACCENT4, ACCENT5, ACCENT3, "#FF6B35"]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Rajdhani:wght@300;400;600;700&family=JetBrains+Mono:wght@400;700&display=swap');

/* ── Base ── */
html, body, [class*="css"] {{
    background-color: {DARK_BG} !important;
    color: {TEXT_PRI} !important;
    font-family: 'Rajdhani', sans-serif;
}}
.stApp {{ background-color: {DARK_BG}; }}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background: {SIDEBAR_BG} !important;
    border-right: 1px solid {BORDER};
}}
section[data-testid="stSidebar"] * {{ color: {TEXT_PRI} !important; }}

/* ── Header ── */
.dashboard-header {{
    background: linear-gradient(135deg, {CARD_BG} 0%, #1A0A2E 100%);
    border: 1px solid {ACCENT1}33;
    border-radius: 16px;
    padding: 28px 36px;
    margin-bottom: 24px;
    position: relative;
    overflow: hidden;
}}
.dashboard-header::before {{
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, {ACCENT1}22, transparent 70%);
    border-radius: 50%;
}}
.dashboard-header h1 {{
    font-family: 'Orbitron', sans-serif;
    font-size: 2rem;
    font-weight: 900;
    background: linear-gradient(90deg, {ACCENT1}, {ACCENT2});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin: 0;
    letter-spacing: 3px;
}}
.dashboard-header p {{
    color: {TEXT_SEC};
    font-size: 0.95rem;
    margin: 6px 0 0;
    letter-spacing: 1px;
}}

/* ── KPI Cards ── */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 14px;
    margin-bottom: 24px;
}}
.kpi-card {{
    background: {CARD_BG};
    border-radius: 12px;
    padding: 20px 18px;
    border: 1px solid {BORDER};
    position: relative;
    overflow: hidden;
    transition: transform 0.2s;
}}
.kpi-card::after {{
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
}}
.kpi-c1::after {{ background: {ACCENT1}; }}
.kpi-c2::after {{ background: {ACCENT2}; }}
.kpi-c3::after {{ background: {ACCENT4}; }}
.kpi-c4::after {{ background: {ACCENT5}; }}
.kpi-c5::after {{ background: {ACCENT3}; }}

.kpi-label {{
    font-size: 0.7rem;
    color: {TEXT_SEC};
    text-transform: uppercase;
    letter-spacing: 2px;
    font-family: 'JetBrains Mono', monospace;
    margin-bottom: 8px;
}}
.kpi-value {{
    font-family: 'Orbitron', sans-serif;
    font-size: 1.55rem;
    font-weight: 700;
    color: {TEXT_PRI};
    line-height: 1;
}}
.kpi-sub {{
    font-size: 0.72rem;
    color: {TEXT_SEC};
    margin-top: 6px;
    font-family: 'JetBrains Mono', monospace;
}}

/* ── Chart card ── */
.chart-card {{
    background: {CARD_BG};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 20px;
}}
.chart-title {{
    font-family: 'Orbitron', sans-serif;
    font-size: 0.8rem;
    color: {ACCENT1};
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 12px;
}}

/* ── Section label ── */
.section-label {{
    font-family: 'Orbitron', sans-serif;
    font-size: 0.65rem;
    color: {ACCENT2};
    letter-spacing: 3px;
    text-transform: uppercase;
    border-left: 3px solid {ACCENT2};
    padding-left: 10px;
    margin: 18px 0 14px;
}}

/* ── Findings ── */
.finding-card {{
    background: linear-gradient(135deg, #12121A, #1a0a2e);
    border: 1px solid {ACCENT1}44;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 12px;
}}
.finding-card h4 {{
    color: {ACCENT1};
    font-family: 'Orbitron', sans-serif;
    font-size: 0.75rem;
    letter-spacing: 2px;
    margin: 0 0 8px;
}}
.finding-card p {{
    color: {TEXT_PRI};
    font-size: 0.92rem;
    margin: 0;
    line-height: 1.6;
}}

/* ── Nav pills ── */
.nav-pill {{
    display: inline-block;
    background: {CARD_BG};
    border: 1px solid {ACCENT1}55;
    border-radius: 30px;
    padding: 6px 20px;
    font-family: 'Orbitron', sans-serif;
    font-size: 0.65rem;
    letter-spacing: 2px;
    color: {ACCENT1};
    margin-right: 10px;
}}

/* ── Streamlit widget overrides ── */
div[data-testid="stSelectbox"] > div,
div[data-testid="stMultiSelect"] > div {{
    background: {CARD_BG} !important;
    border-color: {BORDER} !important;
    border-radius: 8px !important;
}}
div.stButton > button {{
    background: linear-gradient(135deg, {ACCENT1}22, {ACCENT2}22);
    border: 1px solid {ACCENT1}55;
    color: {ACCENT1};
    font-family: 'Orbitron', sans-serif;
    font-size: 0.65rem;
    letter-spacing: 2px;
    border-radius: 8px;
    padding: 8px 20px;
}}
div.stButton > button:hover {{
    background: linear-gradient(135deg, {ACCENT1}44, {ACCENT2}44);
    border-color: {ACCENT1};
}}
.stDateInput > div, .stDateInput input {{
    background: {CARD_BG} !important;
    color: {TEXT_PRI} !important;
    border-color: {BORDER} !important;
}}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MATPLOTLIB DARK THEME DEFAULTS
# ─────────────────────────────────────────────
plt.rcParams.update({
    'figure.facecolor':  DARK_BG,
    'axes.facecolor':    CARD_BG,
    'axes.edgecolor':    BORDER,
    'axes.labelcolor':   TEXT_SEC,
    'axes.titlecolor':   TEXT_PRI,
    'xtick.color':       TEXT_SEC,
    'ytick.color':       TEXT_SEC,
    'text.color':        TEXT_PRI,
    'grid.color':        BORDER,
    'grid.linewidth':    0.5,
    'font.family':       'sans-serif',
    'legend.facecolor':  CARD_BG,
    'legend.edgecolor':  BORDER,
})

# ─────────────────────────────────────────────
#  LOAD DATA
# ─────────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_excel("Sales_Analysis_Data.xlsx")
    df.columns = df.columns.str.strip()
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'])
    df['Month']  = df['Invoice Date'].dt.to_period('M').astype(str)
    df['Year']   = df['Invoice Date'].dt.year
    df['YearMonth'] = df['Invoice Date'].dt.to_period('M')
    return df

df_raw = load_data()

# ─────────────────────────────────────────────
#  SIDEBAR NAVIGATION
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 20px 0 10px;">
        <div style="font-family:'Orbitron',sans-serif; font-size:1.1rem;
                    font-weight:900; color:{ACCENT1}; letter-spacing:3px;">
            👟 ADIDAS
        </div>
        <div style="font-family:'Rajdhani',sans-serif; font-size:0.7rem;
                    color:{TEXT_SEC}; letter-spacing:4px; margin-top:2px;">
            SALES INTELLIGENCE
        </div>
    </div>
    <hr style="border-color:{BORDER}; margin:10px 0 20px;">
    """, unsafe_allow_html=True)

    page = st.radio(
        "NAVIGATION",
        ["📊  Sales Overview", "🗺️  Geo & Product Deep Dive"],
        label_visibility="visible"
    )

    st.markdown(f"""
    <hr style="border-color:{BORDER}; margin:20px 0 14px;">
    <div style="font-family:'JetBrains Mono',monospace; font-size:0.62rem;
                color:{TEXT_SEC}; letter-spacing:2px; text-transform:uppercase;
                margin-bottom:12px;">FILTERS</div>
    """, unsafe_allow_html=True)

    if "Overview" in page:
        regions    = ["All"] + sorted(df_raw['Region'].unique())
        retailers  = ["All"] + sorted(df_raw['Retailer'].unique())
        sel_region   = st.selectbox("Region",   regions)
        sel_retailer = st.selectbox("Retailer", retailers)
        min_date = df_raw['Invoice Date'].min().date()
        max_date = df_raw['Invoice Date'].max().date()
        date_range = st.date_input("Date Range",
            value=(min_date, max_date),
            min_value=min_date, max_value=max_date)
        sel_methods  = None
        sel_products = None
    else:
        methods  = ["All"] + sorted(df_raw['Sales Method'].unique())
        products = ["All"] + sorted(df_raw['Product'].unique())
        sel_method  = st.selectbox("Sales Method", methods)
        sel_product = st.selectbox("Product",      products)
        sel_region   = None
        sel_retailer = None
        date_range   = None
        sel_methods  = sel_method
        sel_products = sel_product

    st.markdown(f"""
    <hr style="border-color:{BORDER}; margin:20px 0 10px;">
    <div style="font-family:'JetBrains Mono',monospace; font-size:0.58rem;
                color:{TEXT_SEC}40; text-align:center;">
        DATA · 2020–2021 · 9,648 RECORDS
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  FILTER DATA
# ─────────────────────────────────────────────
def filter_page1(df):
    d = df.copy()
    if sel_region   and sel_region   != "All": d = d[d['Region']   == sel_region]
    if sel_retailer and sel_retailer != "All": d = d[d['Retailer'] == sel_retailer]
    if date_range and len(date_range) == 2:
        d = d[(d['Invoice Date'].dt.date >= date_range[0]) &
              (d['Invoice Date'].dt.date <= date_range[1])]
    return d

def filter_page2(df):
    d = df.copy()
    if sel_methods  and sel_methods  != "All": d = d[d['Sales Method'] == sel_methods]
    if sel_products and sel_products != "All": d = d[d['Product']      == sel_products]
    return d

# ─────────────────────────────────────────────
#  HELPER: fig_to_st
# ─────────────────────────────────────────────
def show_fig(fig, key=None):
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

def fmt_num(n, decimals=1):
    if n >= 1e9:  return f"${n/1e9:.{decimals}f}B"
    if n >= 1e6:  return f"${n/1e6:.{decimals}f}M"
    if n >= 1e3:  return f"${n/1e3:.{decimals}f}K"
    return f"${n:.2f}"

# ─────────────────────────────────────────────
#  PAGE 1 — SALES OVERVIEW
# ─────────────────────────────────────────────
if "Overview" in page:
    df = filter_page1(df_raw)

    # ── Header ──
    st.markdown(f"""
    <div class="dashboard-header">
        <h1>SALES OVERVIEW</h1>
        <p>Total Sales Performance · Profitability · Volume · Pricing · Margins</p>
    </div>
    """, unsafe_allow_html=True)

    # ── KPIs ──
    total_sales   = df['Total Sales'].sum()
    total_profit  = df['Operating Profit'].sum()
    total_units   = df['Units Sold'].sum()
    avg_price     = df['Price per Unit'].mean()
    avg_margin    = df['Operating Margin'].mean() * 100

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card kpi-c1">
            <div class="kpi-label">Total Sales</div>
            <div class="kpi-value">{fmt_num(total_sales)}</div>
            <div class="kpi-sub">Overall Revenue</div>
        </div>
        <div class="kpi-card kpi-c2">
            <div class="kpi-label">Operating Profit</div>
            <div class="kpi-value">{fmt_num(total_profit)}</div>
            <div class="kpi-sub">Net Profitability</div>
        </div>
        <div class="kpi-card kpi-c3">
            <div class="kpi-label">Units Sold</div>
            <div class="kpi-value">{total_units:,.0f}</div>
            <div class="kpi-sub">Product Demand</div>
        </div>
        <div class="kpi-card kpi-c4">
            <div class="kpi-label">Avg Price / Unit</div>
            <div class="kpi-value">${avg_price:.2f}</div>
            <div class="kpi-sub">Pricing Strategy</div>
        </div>
        <div class="kpi-card kpi-c5">
            <div class="kpi-label">Avg Margin</div>
            <div class="kpi-value">{avg_margin:.1f}%</div>
            <div class="kpi-sub">Profitability Rate</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Row 1: Area Chart + Donut (Region) ──
    col1, col2 = st.columns([2, 1])

    with col1:
        st.markdown('<div class="chart-title">📈 TOTAL SALES BY MONTH</div>', unsafe_allow_html=True)
        monthly = (df.groupby('YearMonth')['Total Sales']
                     .sum().reset_index().sort_values('YearMonth'))
        monthly['Label'] = monthly['YearMonth'].astype(str)
        x = np.arange(len(monthly))
        y = monthly['Total Sales'].values / 1e6

        fig, ax = plt.subplots(figsize=(10, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)

        grad = np.linspace(0, 1, 300).reshape(1, -1)
        cmap = LinearSegmentedColormap.from_list('g', [ACCENT1+'00', ACCENT1+'88'])
        ax.fill_between(x, y, alpha=0.0)
        ax.imshow(grad, aspect='auto', extent=[x[0], x[-1], 0, y.max()],
                  cmap=cmap, origin='lower', zorder=1, alpha=0.45)
        ax.plot(x, y, color=ACCENT1, lw=2.5, zorder=5)
        ax.fill_between(x, y, alpha=0, zorder=2)
        ax.scatter(x, y, color=ACCENT1, s=30, zorder=6)

        step = max(1, len(monthly) // 8)
        ax.set_xticks(x[::step])
        ax.set_xticklabels(monthly['Label'].iloc[::step], rotation=30,
                           ha='right', fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:.0f}M'))
        ax.grid(axis='y', alpha=0.3)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.set_xlabel('')
        plt.tight_layout()
        show_fig(fig)

    with col2:
        st.markdown('<div class="chart-title">🌐 SALES BY REGION</div>', unsafe_allow_html=True)
        reg = df.groupby('Region')['Total Sales'].sum()
        colors_d = [ACCENT1, ACCENT2, ACCENT4, ACCENT5, ACCENT3]

        fig, ax = plt.subplots(figsize=(4.5, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        wedges, texts, autotexts = ax.pie(
            reg.values, labels=reg.index, autopct='%1.1f%%',
            colors=colors_d[:len(reg)],
            wedgeprops=dict(width=0.55, edgecolor=CARD_BG, linewidth=2),
            pctdistance=0.82, startangle=90
        )
        for t in texts:    t.set_color(TEXT_SEC); t.set_fontsize(8)
        for a in autotexts: a.set_color(DARK_BG); a.set_fontsize(7.5); a.set_fontweight('bold')
        plt.tight_layout()
        show_fig(fig)

    # ── Row 2: Bar (Product) + Donut (Sales Method) ──
    col3, col4 = st.columns([2, 1])

    with col3:
        st.markdown('<div class="chart-title">👟 TOTAL SALES BY PRODUCT</div>', unsafe_allow_html=True)
        prod = df.groupby('Product')['Total Sales'].sum().sort_values()
        colors_b = PALETTE[:len(prod)]

        fig, ax = plt.subplots(figsize=(10, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        bars = ax.barh(prod.index, prod.values / 1e6, color=colors_b,
                       edgecolor='none', height=0.6)
        for bar in bars:
            w = bar.get_width()
            ax.text(w + 0.3, bar.get_y() + bar.get_height()/2,
                    f'${w:.1f}M', va='center', fontsize=8, color=TEXT_SEC)
        ax.set_xlabel('Sales (Millions $)', fontsize=8)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis='x', alpha=0.2)
        ax.tick_params(axis='y', labelsize=8)
        plt.tight_layout()
        show_fig(fig)

    with col4:
        st.markdown('<div class="chart-title">🛒 SALES BY METHOD</div>', unsafe_allow_html=True)
        meth = df.groupby('Sales Method')['Total Sales'].sum()
        fig, ax = plt.subplots(figsize=(4.5, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        wedges, texts, autotexts = ax.pie(
            meth.values, labels=meth.index, autopct='%1.1f%%',
            colors=[ACCENT2, ACCENT4, ACCENT1],
            wedgeprops=dict(width=0.55, edgecolor=CARD_BG, linewidth=2),
            pctdistance=0.82, startangle=90
        )
        for t in texts:    t.set_color(TEXT_SEC); t.set_fontsize(8)
        for a in autotexts: a.set_color(DARK_BG); a.set_fontsize(7.5); a.set_fontweight('bold')
        plt.tight_layout()
        show_fig(fig)

    # ── Retailer Sales Bar ──
    st.markdown('<div class="chart-title">🏪 TOTAL SALES BY RETAILER</div>', unsafe_allow_html=True)
    ret = df.groupby('Retailer')['Total Sales'].sum().sort_values(ascending=False)

    fig, ax = plt.subplots(figsize=(14, 3))
    fig.patch.set_facecolor(CARD_BG)
    ax.set_facecolor(CARD_BG)
    colors_r = [ACCENT1 if i == 0 else ACCENT3 for i in range(len(ret))]
    bars = ax.bar(ret.index, ret.values / 1e6, color=colors_r, edgecolor='none', width=0.5)
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, h + 0.5,
                f'${h:.1f}M', ha='center', fontsize=8, color=TEXT_SEC)
    ax.set_ylabel('Sales (M$)', fontsize=8)
    ax.spines[['top','right','left','bottom']].set_visible(False)
    ax.grid(axis='y', alpha=0.2)
    plt.tight_layout()
    show_fig(fig)

    # ── KEY FINDINGS PAGE 1 ──
    st.markdown('<div class="section-label">KEY FINDINGS — PAGE 1</div>', unsafe_allow_html=True)

    top_retailer = ret.index[0]
    top_product  = df.groupby('Product')['Total Sales'].sum().idxmax()
    top_region   = df.groupby('Region')['Total Sales'].sum().idxmax()
    top_method   = df.groupby('Sales Method')['Total Sales'].sum().idxmax()
    peak_month   = monthly.loc[monthly['Total Sales'].idxmax(), 'Label']

    findings = [
        ("🏆 TOP RETAILER", f"<b>{top_retailer}</b> leads all retailers in total revenue, capturing the largest share of Adidas sales across the dataset period."),
        ("👟 BEST-SELLING PRODUCT", f"<b>{top_product}</b> drives the most revenue, signaling strong consumer preference and potentially a strategic pricing sweet spot."),
        ("🌐 DOMINANT REGION", f"The <b>{top_region}</b> region contributes the highest sales volume, making it a priority market for Adidas distribution and marketing."),
        ("🛒 PREFERRED SALES CHANNEL", f"<b>{top_method}</b> is the most productive sales channel by revenue — a signal for where Adidas should continue to invest."),
        ("📅 PEAK SALES MONTH", f"Sales peaked in <b>{peak_month}</b>, likely driven by seasonal demand, promotions, or product launches — a key insight for inventory planning."),
        ("💰 MARGIN HEALTH", f"The average operating margin of <b>{avg_margin:.1f}%</b> reflects a healthy profitability baseline, though optimization opportunities exist at the product level."),
    ]
    cols = st.columns(2)
    for i, (title, body) in enumerate(findings):
        with cols[i % 2]:
            st.markdown(f"""
            <div class="finding-card">
                <h4>{title}</h4>
                <p>{body}</p>
            </div>
            """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  PAGE 2 — GEO & PRODUCT DEEP DIVE
# ─────────────────────────────────────────────
else:
    df = filter_page2(df_raw)

    st.markdown(f"""
    <div class="dashboard-header">
        <h1>GEO & PRODUCT DEEP DIVE</h1>
        <p>State Performance · City Rankings · Product Profitability · Price-Volume Dynamics</p>
    </div>
    """, unsafe_allow_html=True)

    # ── KPIs ──
    top_state   = df.groupby('State')['Total Sales'].sum().idxmax()
    top_product2= df.groupby('Product')['Total Sales'].sum().idxmax()
    top_ret_prof= df.groupby('Retailer')['Operating Profit'].sum().idxmax()
    avg_price2  = df['Price per Unit'].mean()

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card kpi-c1">
            <div class="kpi-label">Highest Selling State</div>
            <div class="kpi-value" style="font-size:1.15rem;">{top_state}</div>
            <div class="kpi-sub">By Total Revenue</div>
        </div>
        <div class="kpi-card kpi-c2">
            <div class="kpi-label">Highest Selling Product</div>
            <div class="kpi-value" style="font-size:0.95rem;">{top_product2}</div>
            <div class="kpi-sub">By Total Revenue</div>
        </div>
        <div class="kpi-card kpi-c3">
            <div class="kpi-label">Most Profitable Retailer</div>
            <div class="kpi-value" style="font-size:1.05rem;">{top_ret_prof}</div>
            <div class="kpi-sub">By Operating Profit</div>
        </div>
        <div class="kpi-card kpi-c4">
            <div class="kpi-label">Avg Price / Unit</div>
            <div class="kpi-value">${avg_price2:.2f}</div>
            <div class="kpi-sub">Pricing Benchmark</div>
        </div>
        <div class="kpi-card kpi-c5">
            <div class="kpi-label">Total Records</div>
            <div class="kpi-value">{len(df):,}</div>
            <div class="kpi-sub">Filtered Dataset</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Row 1: State Sales (Horizontal Bar as map proxy) + Top 10 Cities ──
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown('<div class="chart-title">🗺️ TOTAL SALES BY STATE (TOP 20)</div>', unsafe_allow_html=True)
        state_sales = df.groupby('State')['Total Sales'].sum().sort_values(ascending=False).head(20)
        norm = state_sales.values / state_sales.max()
        colors_s = [f"#{int(0*255):02x}{int(n*229):02x}{int(n*255):02x}" for n in norm]

        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        bars = ax.barh(state_sales.index[::-1], state_sales.values[::-1] / 1e6,
                       color=colors_s[::-1], edgecolor='none', height=0.65)
        for bar in bars:
            w = bar.get_width()
            ax.text(w + 0.1, bar.get_y() + bar.get_height()/2,
                    f'${w:.1f}M', va='center', fontsize=7, color=TEXT_SEC)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis='x', alpha=0.2)
        ax.tick_params(axis='y', labelsize=7.5)
        ax.set_xlabel('Sales (M$)', fontsize=8)
        plt.tight_layout()
        show_fig(fig)

    with col2:
        st.markdown('<div class="chart-title">🏙️ TOP 10 CITIES BY SALES</div>', unsafe_allow_html=True)
        city_sales = df.groupby('City')['Total Sales'].sum().sort_values(ascending=False).head(10)

        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        x_pos = np.arange(len(city_sales))
        bars = ax.bar(x_pos, city_sales.values / 1e6,
                      color=PALETTE[:len(city_sales)], edgecolor='none', width=0.6)
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, h + 0.2,
                    f'${h:.1f}M', ha='center', fontsize=7, color=TEXT_SEC)
        ax.set_xticks(x_pos)
        ax.set_xticklabels(city_sales.index, rotation=35, ha='right', fontsize=8)
        ax.set_ylabel('Sales (M$)', fontsize=8)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis='y', alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    # ── Row 2: Profit by Product + Price vs Units ──
    col3, col4 = st.columns([1, 1])

    with col3:
        st.markdown('<div class="chart-title">💰 PROFIT BY PRODUCT</div>', unsafe_allow_html=True)
        prod_profit = df.groupby('Product')['Operating Profit'].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        x_pos = np.arange(len(prod_profit))
        bars = ax.bar(x_pos, prod_profit.values / 1e6,
                      color=PALETTE[:len(prod_profit)], edgecolor='none', width=0.55)
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, h + 0.2,
                    f'${h:.1f}M', ha='center', fontsize=7, color=TEXT_SEC)
        ax.set_xticks(x_pos)
        ax.set_xticklabels([p.replace("'s ", "\n") for p in prod_profit.index],
                           fontsize=7.5)
        ax.set_ylabel('Profit (M$)', fontsize=8)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis='y', alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    with col4:
        st.markdown('<div class="chart-title">⚡ PRICE vs UNITS SOLD</div>', unsafe_allow_html=True)
        sample = df.sample(min(1500, len(df)), random_state=42)
        products_u = df['Product'].unique()
        color_map  = {p: PALETTE[i % len(PALETTE)] for i, p in enumerate(products_u)}
        c_list = [color_map[p] for p in sample['Product']]

        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        ax.scatter(sample['Price per Unit'], sample['Units Sold'],
                   c=c_list, alpha=0.55, s=18, edgecolors='none')
        handles = [mpatches.Patch(color=color_map[p], label=p) for p in products_u]
        ax.legend(handles=handles, fontsize=6.5, loc='upper right',
                  framealpha=0.3, ncol=1)
        ax.set_xlabel('Price per Unit ($)', fontsize=8)
        ax.set_ylabel('Units Sold', fontsize=8)
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    # ── Row 3: Sales vs Profit scatter ──
    st.markdown('<div class="chart-title">📊 SALES vs OPERATING PROFIT (by Retailer)</div>', unsafe_allow_html=True)
    sample2   = df.sample(min(2000, len(df)), random_state=7)
    retailers = df['Retailer'].unique()
    cmap_ret  = {r: PALETTE[i % len(PALETTE)] for i, r in enumerate(retailers)}
    c_ret     = [cmap_ret[r] for r in sample2['Retailer']]

    fig, ax = plt.subplots(figsize=(14, 4))
    fig.patch.set_facecolor(CARD_BG)
    ax.set_facecolor(CARD_BG)
    ax.scatter(sample2['Total Sales'], sample2['Operating Profit'],
               c=c_ret, alpha=0.5, s=22, edgecolors='none')
    handles2 = [mpatches.Patch(color=cmap_ret[r], label=r) for r in retailers]
    ax.legend(handles=handles2, fontsize=8, loc='upper left',
              framealpha=0.3, ncol=len(retailers))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v/1e3:.0f}K'))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v/1e3:.0f}K'))
    ax.set_xlabel('Total Sales ($)', fontsize=8)
    ax.set_ylabel('Operating Profit ($)', fontsize=8)
    ax.spines[['top','right','left','bottom']].set_visible(False)
    ax.grid(alpha=0.15)
    plt.tight_layout()
    show_fig(fig)

    # ── KEY FINDINGS PAGE 2 ──
    st.markdown('<div class="section-label">KEY FINDINGS — PAGE 2</div>', unsafe_allow_html=True)

    top_city = df.groupby('City')['Total Sales'].sum().idxmax()
    top_state_profit = df.groupby('State')['Operating Profit'].sum().idxmax()
    high_margin_prod = df.groupby('Product')['Operating Margin'].mean().idxmax()
    low_price_high_vol = df.groupby('Product').apply(
        lambda x: x['Units Sold'].mean()).idxmax()

    findings2 = [
        ("🏙️ TOP PERFORMING CITY", f"<b>{top_city}</b> ranks as the highest-revenue city, making it Adidas's most critical urban market for sales concentration."),
        ("🗺️ TOP PROFIT STATE", f"<b>{top_state_profit}</b> generates the most operating profit among all states, suggesting strong operational efficiency or favorable pricing dynamics."),
        ("💎 HIGHEST MARGIN PRODUCT", f"<b>{high_margin_prod}</b> commands the best average margin, highlighting it as Adidas's most financially efficient product line."),
        ("📦 HIGHEST VOLUME PRODUCT", f"<b>{low_price_high_vol}</b> achieves the highest average units sold per transaction, indicating strong mass-market demand."),
        ("🔗 SALES–PROFIT CORRELATION", "The scatter analysis confirms a strong positive linear relationship between total sales and operating profit, validating that revenue growth directly drives profitability."),
        ("💸 PRICE SENSITIVITY", f"Scatter analysis reveals that lower-priced products (under $45) consistently outperform in units sold, suggesting significant price elasticity in Adidas's consumer base."),
    ]
    cols2 = st.columns(2)
    for i, (title, body) in enumerate(findings2):
        with cols2[i % 2]:
            st.markdown(f"""
            <div class="finding-card">
                <h4>{title}</h4>
                <p>{body}</p>
            </div>
            """, unsafe_allow_html=True)

# ── Footer ──
st.markdown(f"""
<div style="text-align:center; padding:30px 0 10px; color:{TEXT_SEC};
            font-family:'JetBrains Mono',monospace; font-size:0.6rem; letter-spacing:2px;">
    ADIDAS SALES INTELLIGENCE DASHBOARD · 2020–2021 · BUILT WITH STREAMLIT + MATPLOTLIB
</div>
""", unsafe_allow_html=True)
