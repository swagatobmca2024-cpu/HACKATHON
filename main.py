import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.colors import LinearSegmentedColormap
import matplotlib.ticker as mticker
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Adidas Sales Dashboard",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS  (dark theme)
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow:wght@300;400;600;700;800&family=Barlow+Condensed:wght@500;700;900&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'Barlow', sans-serif;
    background: #080b14;
    color: #dde3f0;
}
.stApp { background: #080b14; }

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #0d1120;
    border-right: 1px solid #1a2240;
    padding-top: 0;
}
[data-testid="stSidebar"] * { color: #dde3f0 !important; }

/* ── Remove default Streamlit chrome ── */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding: 1.5rem 2rem 3rem 2rem; }

/* ── Navigation pills ── */
.nav-wrap {
    display: flex;
    gap: 10px;
    margin-bottom: 28px;
    padding: 6px;
    background: #0d1120;
    border-radius: 14px;
    border: 1px solid #1a2240;
    width: fit-content;
}
.nav-btn {
    padding: 10px 28px;
    border-radius: 10px;
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 15px;
    font-weight: 700;
    letter-spacing: 1px;
    text-transform: uppercase;
    cursor: pointer;
    border: none;
    transition: all 0.25s ease;
    text-decoration: none;
    display: inline-block;
}
.nav-active {
    background: linear-gradient(135deg, #00e5ff, #0072ff);
    color: #080b14 !important;
    box-shadow: 0 4px 20px #0072ff55;
}
.nav-inactive {
    background: transparent;
    color: #7a8ab0 !important;
}
.nav-inactive:hover { background: #1a2240; color: #dde3f0 !important; }

/* ── Page title ── */
.page-header {
    margin-bottom: 28px;
    padding-bottom: 16px;
    border-bottom: 1px solid #1a2240;
}
.page-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 36px;
    font-weight: 900;
    letter-spacing: 2px;
    text-transform: uppercase;
    background: linear-gradient(135deg, #00e5ff, #0072ff);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin: 0;
    line-height: 1;
}
.page-subtitle {
    font-size: 13px;
    color: #4a5a80;
    margin-top: 6px;
    letter-spacing: 0.5px;
}

/* ── KPI Cards ── */
.kpi-grid {
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 14px;
    margin-bottom: 24px;
}
.kpi-card {
    background: #0d1120;
    border: 1px solid #1a2240;
    border-radius: 14px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s;
}
.kpi-card:hover { border-color: #0072ff66; }
.kpi-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 3px;
    border-radius: 14px 14px 0 0;
}
.kpi-blue::before   { background: linear-gradient(90deg, #0072ff, #00e5ff); }
.kpi-green::before  { background: linear-gradient(90deg, #00cc88, #00ffb3); }
.kpi-purple::before { background: linear-gradient(90deg, #8b5cf6, #c084fc); }
.kpi-amber::before  { background: linear-gradient(90deg, #f59e0b, #fcd34d); }
.kpi-rose::before   { background: linear-gradient(90deg, #f43f5e, #fb7185); }

.kpi-icon { font-size: 22px; margin-bottom: 10px; }
.kpi-val {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 28px;
    font-weight: 900;
    letter-spacing: 0.5px;
    line-height: 1;
    margin-bottom: 4px;
}
.kpi-lbl {
    font-size: 11px;
    color: #4a5a80;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    font-weight: 600;
}
.kpi-delta {
    font-size: 12px;
    margin-top: 8px;
    font-family: 'JetBrains Mono', monospace;
}
.delta-up   { color: #00cc88; }
.delta-down { color: #f43f5e; }

/* ── Chart containers ── */
.chart-card {
    background: #0d1120;
    border: 1px solid #1a2240;
    border-radius: 14px;
    padding: 20px;
    margin-bottom: 18px;
}
.chart-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 16px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #7a9fff;
    margin-bottom: 4px;
}
.chart-subtitle {
    font-size: 12px;
    color: #3a4a6a;
    margin-bottom: 14px;
}

/* ── Insight box ── */
.insight-box {
    background: #0d1120;
    border: 1px solid #1a2240;
    border-left: 4px solid #00e5ff;
    border-radius: 10px;
    padding: 16px 20px;
    margin-bottom: 12px;
    font-size: 14px;
    line-height: 1.6;
}
.insight-title {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 13px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #00e5ff;
    margin-bottom: 6px;
}

/* ── Section divider ── */
.section-divider {
    height: 1px;
    background: linear-gradient(90deg, #1a2240, transparent);
    margin: 24px 0;
}

/* ── Sidebar filter labels ── */
.filter-label {
    font-family: 'Barlow Condensed', sans-serif;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 1.5px;
    color: #4a5a80 !important;
    margin-bottom: 6px;
    margin-top: 14px;
}

/* ── Multiselect & date styling ── */
[data-testid="stMultiSelect"] > div,
[data-testid="stDateInput"] > div {
    background: #131929 !important;
    border: 1px solid #1a2240 !important;
    border-radius: 8px !important;
}

/* ── Streamlit overrides ── */
.stSlider > div { color: #00e5ff !important; }
.stDateInput label, .stMultiSelect label, .stSelectbox label {
    color: #7a8ab0 !important;
    font-size: 12px !important;
    font-weight: 600 !important;
    text-transform: uppercase !important;
    letter-spacing: 1px !important;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 4px; height: 4px; }
::-webkit-scrollbar-track { background: #0d1120; }
::-webkit-scrollbar-thumb { background: #1a2240; border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# MATPLOTLIB DARK THEME
# ─────────────────────────────────────────────────────────────────────────────
DARK_BG    = "#0d1120"
GRID_COLOR = "#1a2240"
TEXT_COLOR = "#dde3f0"
ACCENT     = "#00e5ff"
ACCENT2    = "#0072ff"
PALETTE    = ["#0072ff","#00e5ff","#00cc88","#8b5cf6","#f59e0b","#f43f5e",
              "#34d399","#60a5fa","#a78bfa","#fb923c"]

plt.rcParams.update({
    "figure.facecolor":  DARK_BG,
    "axes.facecolor":    DARK_BG,
    "axes.edgecolor":    GRID_COLOR,
    "axes.labelcolor":   TEXT_COLOR,
    "axes.titlecolor":   TEXT_COLOR,
    "xtick.color":       TEXT_COLOR,
    "ytick.color":       TEXT_COLOR,
    "grid.color":        GRID_COLOR,
    "grid.linewidth":    0.6,
    "text.color":        TEXT_COLOR,
    "font.family":       "DejaVu Sans",
    "figure.dpi":        120,
    "savefig.facecolor": DARK_BG,
    "savefig.dpi":       120,
})

def fig_style(ax, title="", xlabel="", ylabel="", grid_axis="y"):
    ax.set_title(title, fontsize=13, fontweight="bold", color=TEXT_COLOR, pad=12)
    ax.set_xlabel(xlabel, fontsize=10, color="#4a5a80", labelpad=8)
    ax.set_ylabel(ylabel, fontsize=10, color="#4a5a80", labelpad=8)
    ax.tick_params(axis='both', labelsize=9, colors=TEXT_COLOR)
    ax.spines[["top","right","left","bottom"]].set_visible(False)
    if grid_axis:
        ax.grid(axis=grid_axis, color=GRID_COLOR, linewidth=0.6, linestyle="--")
    ax.set_axisbelow(True)

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data
def load_data():
    df = pd.read_excel("Sales_Analysis_Data.xlsx")
    df["Invoice Date"] = pd.to_datetime(df["Invoice Date"])
    df["Month"]        = df["Invoice Date"].dt.to_period("M").dt.to_timestamp()
    df["YearMonth"]    = df["Invoice Date"].dt.strftime("%b %Y")
    df["Year"]         = df["Invoice Date"].dt.year
    return df

df_master = load_data()

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE — page navigation
# ─────────────────────────────────────────────────────────────────────────────
if "page" not in st.session_state:
    st.session_state["page"] = "page1"

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR LOGO + NAV
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='padding:20px 10px 10px 10px;'>
        <div style='font-family:"Barlow Condensed",sans-serif;font-size:28px;font-weight:900;
                    background:linear-gradient(135deg,#00e5ff,#0072ff);
                    -webkit-background-clip:text;-webkit-text-fill-color:transparent;
                    letter-spacing:3px;'>ADIDAS</div>
        <div style='font-size:11px;color:#3a4a6a;letter-spacing:2px;
                    text-transform:uppercase;margin-top:2px;'>Sales Intelligence</div>
    </div>
    <hr style='border:none;border-top:1px solid #1a2240;margin:10px 0 18px 0;'>
    """, unsafe_allow_html=True)

    p1 = st.button("📊  Page 1 — Sales Overview",  use_container_width=True)
    p2 = st.button("🗺️   Page 2 — Regional & Retail", use_container_width=True)

    if p1: st.session_state["page"] = "page1"
    if p2: st.session_state["page"] = "page2"

    st.markdown("<hr style='border:none;border-top:1px solid #1a2240;margin:18px 0;'>", unsafe_allow_html=True)

    # ── PAGE 1 filters ──
    if st.session_state["page"] == "page1":
        st.markdown("<div class='filter-label'>🌍 Region</div>", unsafe_allow_html=True)
        regions = st.multiselect("Region", options=sorted(df_master["Region"].unique()),
                                 default=sorted(df_master["Region"].unique()), label_visibility="collapsed")

        st.markdown("<div class='filter-label'>🏪 Retailer</div>", unsafe_allow_html=True)
        retailers = st.multiselect("Retailer", options=sorted(df_master["Retailer"].unique()),
                                   default=sorted(df_master["Retailer"].unique()), label_visibility="collapsed")

        st.markdown("<div class='filter-label'>📅 Date Range</div>", unsafe_allow_html=True)
        min_date = df_master["Invoice Date"].min().date()
        max_date = df_master["Invoice Date"].max().date()
        date_range = st.date_input("Date", value=(min_date, max_date),
                                   min_value=min_date, max_value=max_date,
                                   label_visibility="collapsed")

    # ── PAGE 2 filters ──
    else:
        st.markdown("<div class='filter-label'>🛒 Sales Method</div>", unsafe_allow_html=True)
        methods = st.multiselect("Method", options=sorted(df_master["Sales Method"].unique()),
                                 default=sorted(df_master["Sales Method"].unique()), label_visibility="collapsed")

        st.markdown("<div class='filter-label'>👟 Product</div>", unsafe_allow_html=True)
        products = st.multiselect("Product", options=sorted(df_master["Product"].unique()),
                                  default=sorted(df_master["Product"].unique()), label_visibility="collapsed")

# ─────────────────────────────────────────────────────────────────────────────
# FILTER DATA
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state["page"] == "page1":
    df = df_master[df_master["Region"].isin(regions) & df_master["Retailer"].isin(retailers)].copy()
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        df = df[(df["Invoice Date"].dt.date >= date_range[0]) &
                (df["Invoice Date"].dt.date <= date_range[1])]
else:
    df = df_master[df_master["Sales Method"].isin(methods) & df_master["Product"].isin(products)].copy()

# ═════════════════════════════════════════════════════════════════════════════
# ███████████████████████  PAGE 1  ███████████████████████████████████████████
# ═════════════════════════════════════════════════════════════════════════════
if st.session_state["page"] == "page1":

    st.markdown("""
    <div class="page-header">
        <div class="page-title">Sales Overview</div>
        <div class="page-subtitle">Adidas United States · 2020–2021 · All figures in USD</div>
    </div>
    """, unsafe_allow_html=True)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    total_sales   = df["Total Sales"].sum()
    total_profit  = df["Operating Profit"].sum()
    total_units   = df["Units Sold"].sum()
    avg_price     = df["Price per Unit"].mean()
    avg_margin    = df["Operating Margin"].mean()

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card kpi-blue">
            <div class="kpi-icon">💰</div>
            <div class="kpi-val">${total_sales/1e6:.1f}M</div>
            <div class="kpi-lbl">Total Sales</div>
            <div class="kpi-delta delta-up">↑ All Regions Combined</div>
        </div>
        <div class="kpi-card kpi-green">
            <div class="kpi-icon">📈</div>
            <div class="kpi-val">${total_profit/1e6:.1f}M</div>
            <div class="kpi-lbl">Total Profit</div>
            <div class="kpi-delta delta-up">↑ Operating Profit</div>
        </div>
        <div class="kpi-card kpi-purple">
            <div class="kpi-icon">📦</div>
            <div class="kpi-val">{total_units/1e6:.2f}M</div>
            <div class="kpi-lbl">Units Sold</div>
            <div class="kpi-delta delta-up">↑ Total Volume</div>
        </div>
        <div class="kpi-card kpi-amber">
            <div class="kpi-icon">🏷️</div>
            <div class="kpi-val">${avg_price:.2f}</div>
            <div class="kpi-lbl">Avg Price / Unit</div>
            <div class="kpi-delta delta-up">↑ Pricing Strategy</div>
        </div>
        <div class="kpi-card kpi-rose">
            <div class="kpi-icon">📊</div>
            <div class="kpi-val">{avg_margin*100:.1f}%</div>
            <div class="kpi-lbl">Avg Margin</div>
            <div class="kpi-delta delta-up">↑ Profitability</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── ROW 1: Area Chart + Donut (Region) ────────────────────────────────────
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Total Sales by Month</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Monthly revenue trend · Area chart</div>', unsafe_allow_html=True)

        monthly = df.groupby("Month")["Total Sales"].sum().reset_index().sort_values("Month")

        fig, ax = plt.subplots(figsize=(9, 3.8))
        x = np.arange(len(monthly))
        y = monthly["Total Sales"].values

        # Gradient fill
        ax.fill_between(x, y, alpha=0.18, color=ACCENT2)
        ax.fill_between(x, y, alpha=0.08, color=ACCENT)
        ax.plot(x, y, color=ACCENT, linewidth=2.2, zorder=5)
        ax.scatter(x, y, color=ACCENT, s=28, zorder=6, linewidth=0)

        # Peak annotation
        peak_idx = np.argmax(y)
        ax.annotate(f"${y[peak_idx]/1e6:.1f}M",
                    xy=(x[peak_idx], y[peak_idx]),
                    xytext=(x[peak_idx], y[peak_idx]*1.06),
                    fontsize=8, color=ACCENT, ha='center', fontweight='bold')

        ax.set_xticks(x[::2])
        ax.set_xticklabels(monthly["Month"].dt.strftime("%b %y").values[::2], rotation=30, ha='right', fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, ylabel="Sales (USD)")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Sales by Region</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Contribution share · Donut chart</div>', unsafe_allow_html=True)

        region_sales = df.groupby("Region")["Total Sales"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(5, 4))
        colors  = PALETTE[:len(region_sales)]
        wedges, texts, autotexts = ax.pie(
            region_sales.values, labels=region_sales.index,
            autopct="%1.1f%%", startangle=90,
            colors=colors, pctdistance=0.78,
            wedgeprops=dict(width=0.55, edgecolor=DARK_BG, linewidth=2)
        )
        for t in texts:      t.set(color=TEXT_COLOR, fontsize=8)
        for a in autotexts:  a.set(color=DARK_BG, fontsize=7.5, fontweight="bold")
        ax.set_facecolor(DARK_BG)
        fig.patch.set_facecolor(DARK_BG)

        # Center text
        top_region = region_sales.index[0]
        ax.text(0, 0, f"Top\n{top_region}", ha='center', va='center',
                fontsize=8, color=ACCENT, fontweight='bold')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── ROW 2: Bar (Product) + Bar (Retailer) ────────────────────────────────
    col3, col4 = st.columns(2)

    with col3:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Total Sales by Product</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Revenue breakdown across product categories</div>', unsafe_allow_html=True)

        prod_sales = df.groupby("Product")["Total Sales"].sum().sort_values()

        fig, ax = plt.subplots(figsize=(6, 3.8))
        bars = ax.barh(
            [p.replace("'", "'") for p in prod_sales.index],
            prod_sales.values,
            color=[PALETTE[i % len(PALETTE)] for i in range(len(prod_sales))],
            height=0.62, edgecolor="none"
        )
        for bar in bars:
            w = bar.get_width()
            ax.text(w * 1.01, bar.get_y() + bar.get_height()/2,
                    f"${w/1e6:.1f}M", va='center', fontsize=8, color=TEXT_COLOR)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, grid_axis="x")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col4:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Total Sales by Retailer</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Revenue performance across retail partners</div>', unsafe_allow_html=True)

        ret_sales = df.groupby("Retailer")["Total Sales"].sum().sort_values()

        fig, ax = plt.subplots(figsize=(6, 3.8))
        bars = ax.barh(
            ret_sales.index, ret_sales.values,
            color=[PALETTE[(i+3) % len(PALETTE)] for i in range(len(ret_sales))],
            height=0.62, edgecolor="none"
        )
        for bar in bars:
            w = bar.get_width()
            ax.text(w * 1.01, bar.get_y() + bar.get_height()/2,
                    f"${w/1e6:.1f}M", va='center', fontsize=8, color=TEXT_COLOR)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, grid_axis="x")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── ROW 3: Donut (Sales Method) + Margin by Product ───────────────────────
    col5, col6 = st.columns([2, 3])

    with col5:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Sales by Method</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Channel distribution · Donut chart</div>', unsafe_allow_html=True)

        method_sales = df.groupby("Sales Method")["Total Sales"].sum()

        fig, ax = plt.subplots(figsize=(4.5, 4))
        mcolors  = ["#0072ff", "#00cc88", "#f59e0b"]
        wedges, texts, autotexts = ax.pie(
            method_sales.values, labels=method_sales.index,
            autopct="%1.1f%%", startangle=90,
            colors=mcolors, pctdistance=0.78,
            wedgeprops=dict(width=0.52, edgecolor=DARK_BG, linewidth=2)
        )
        for t in texts:      t.set(color=TEXT_COLOR, fontsize=9)
        for a in autotexts:  a.set(color=DARK_BG, fontsize=8, fontweight="bold")
        fig.patch.set_facecolor(DARK_BG)
        ax.set_facecolor(DARK_BG)
        top_method = method_sales.idxmax()
        ax.text(0, 0, f"Top\n{top_method}", ha='center', va='center',
                fontsize=8, color=ACCENT, fontweight='bold')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col6:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Avg Operating Margin by Product</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Profitability efficiency across categories</div>', unsafe_allow_html=True)

        margin_prod = df.groupby("Product")["Operating Margin"].mean().sort_values()

        fig, ax = plt.subplots(figsize=(7, 3.8))
        bar_colors = [PALETTE[i % len(PALETTE)] for i in range(len(margin_prod))]
        bars = ax.barh(margin_prod.index, margin_prod.values * 100,
                       color=bar_colors, height=0.62, edgecolor="none")
        for bar in bars:
            w = bar.get_width()
            ax.text(w + 0.3, bar.get_y() + bar.get_height()/2,
                    f"{w:.1f}%", va='center', fontsize=8.5, color=TEXT_COLOR)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
        ax.set_xlim(0, margin_prod.values.max()*100 * 1.15)
        fig_style(ax, grid_axis="x", xlabel="Operating Margin (%)")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── KEY FINDINGS PAGE 1 ────────────────────────────────────────────────────
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div style='font-family:"Barlow Condensed",sans-serif;font-size:18px;font-weight:700;
                text-transform:uppercase;letter-spacing:2px;color:#7a9fff;margin-bottom:14px;'>
        🔍 Key Findings — Sales Overview
    </div>
    """, unsafe_allow_html=True)

    top_region_val = df.groupby("Region")["Total Sales"].sum().idxmax()
    top_product    = df.groupby("Product")["Total Sales"].sum().idxmax()
    top_retailer   = df.groupby("Retailer")["Total Sales"].sum().idxmax()
    top_method     = df.groupby("Sales Method")["Total Sales"].sum().idxmax()
    peak_month     = df.groupby("Month")["Total Sales"].sum().idxmax().strftime("%B %Y")
    best_margin_p  = df.groupby("Product")["Operating Margin"].mean().idxmax()

    findings = [
        ("Revenue Leadership", f"<b>{top_region_val}</b> region leads all regions in total sales, driven by high population density and retail presence."),
        ("Top Product", f"<b>{top_product}</b> is the highest-selling product category, reflecting strong consumer demand for street footwear."),
        ("Retail Powerhouse", f"<b>{top_retailer}</b> is the #1 retail partner by total sales volume, underscoring the importance of specialty retail channels."),
        ("Sales Channel", f"<b>{top_method}</b> drives the majority of revenue, though online growth signals a shift in consumer buying behavior."),
        ("Peak Period", f"Sales peaked in <b>{peak_month}</b>, suggesting seasonal demand — likely linked to back-to-school or holiday shopping cycles."),
        ("Margin Leader", f"<b>{best_margin_p}</b> generates the highest average operating margin, making it the most profitable product line per dollar of revenue."),
    ]

    c_a, c_b = st.columns(2)
    for i, (title, body) in enumerate(findings):
        target = c_a if i % 2 == 0 else c_b
        with target:
            st.markdown(f"""
            <div class="insight-box">
                <div class="insight-title">{title}</div>
                {body}
            </div>""", unsafe_allow_html=True)


# ═════════════════════════════════════════════════════════════════════════════
# ███████████████████████  PAGE 2  ███████████████████████████████████████████
# ═════════════════════════════════════════════════════════════════════════════
else:
    st.markdown("""
    <div class="page-header">
        <div class="page-title">Regional & Retail Deep Dive</div>
        <div class="page-subtitle">State · City · Product · Retailer breakdowns · 2020–2021</div>
    </div>
    """, unsafe_allow_html=True)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    top_state    = df.groupby("State")["Total Sales"].sum().idxmax()
    top_state_v  = df.groupby("State")["Total Sales"].sum().max()
    top_product2 = df.groupby("Product")["Total Sales"].sum().idxmax()
    top_product2_v = df.groupby("Product")["Total Sales"].sum().max()
    top_retailer2  = df.groupby("Retailer")["Operating Profit"].sum().idxmax()
    top_ret_profit = df.groupby("Retailer")["Operating Profit"].sum().max()
    avg_price2   = df["Price per Unit"].mean()

    st.markdown(f"""
    <div class="kpi-grid">
        <div class="kpi-card kpi-blue">
            <div class="kpi-icon">🗺️</div>
            <div class="kpi-val">{top_state}</div>
            <div class="kpi-lbl">Highest Selling State</div>
            <div class="kpi-delta delta-up">${top_state_v/1e6:.1f}M in sales</div>
        </div>
        <div class="kpi-card kpi-green">
            <div class="kpi-icon">👟</div>
            <div class="kpi-val" style="font-size:18px;">{top_product2.replace("Men's","M").replace("Women's","W")}</div>
            <div class="kpi-lbl">Highest Selling Product</div>
            <div class="kpi-delta delta-up">${top_product2_v/1e6:.1f}M revenue</div>
        </div>
        <div class="kpi-card kpi-purple">
            <div class="kpi-icon">🏪</div>
            <div class="kpi-val" style="font-size:20px;">{top_retailer2}</div>
            <div class="kpi-lbl">Most Profitable Retailer</div>
            <div class="kpi-delta delta-up">${top_ret_profit/1e6:.1f}M profit</div>
        </div>
        <div class="kpi-card kpi-amber">
            <div class="kpi-icon">🏷️</div>
            <div class="kpi-val">${avg_price2:.2f}</div>
            <div class="kpi-lbl">Avg Price per Unit</div>
            <div class="kpi-delta delta-up">Blended average</div>
        </div>
        <div class="kpi-card kpi-rose">
            <div class="kpi-icon">💹</div>
            <div class="kpi-val">${df["Operating Profit"].sum()/1e6:.1f}M</div>
            <div class="kpi-lbl">Total Operating Profit</div>
            <div class="kpi-delta delta-up">Filtered selection</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── ROW 1: State bar (horizontal, top 15) + Top 10 Cities ─────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Total Sales by State (Top 15)</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Filled map proxy · horizontal bar ranking</div>', unsafe_allow_html=True)

        state_sales = df.groupby("State")["Total Sales"].sum().nlargest(15).sort_values()

        fig, ax = plt.subplots(figsize=(6, 5))
        # Color gradient by rank
        norm_vals = (state_sales.values - state_sales.values.min()) / (state_sales.values.max() - state_sales.values.min())
        bar_colors = [plt.cm.cool(v * 0.75 + 0.1) for v in norm_vals]
        bars = ax.barh(state_sales.index, state_sales.values, color=bar_colors, height=0.72, edgecolor="none")
        for bar in bars:
            w = bar.get_width()
            ax.text(w * 1.01, bar.get_y() + bar.get_height()/2,
                    f"${w/1e6:.1f}M", va='center', fontsize=7.5, color=TEXT_COLOR)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, grid_axis="x")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Top 10 Cities by Sales</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Urban revenue hotspots · bar chart</div>', unsafe_allow_html=True)

        city_sales = df.groupby("City")["Total Sales"].sum().nlargest(10).sort_values()

        fig, ax = plt.subplots(figsize=(6, 5))
        norm_vals2 = (city_sales.values - city_sales.values.min()) / (city_sales.values.max() - city_sales.values.min())
        bar_colors2 = [plt.cm.plasma(v * 0.6 + 0.25) for v in norm_vals2]
        bars = ax.barh(city_sales.index, city_sales.values, color=bar_colors2, height=0.72, edgecolor="none")
        for bar in bars:
            w = bar.get_width()
            ax.text(w * 1.01, bar.get_y() + bar.get_height()/2,
                    f"${w/1e6:.1f}M", va='center', fontsize=7.5, color=TEXT_COLOR)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, grid_axis="x")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── ROW 2: Profit by Product (column) + Price vs Units (scatter) ──────────
    col3, col4 = st.columns(2)

    with col3:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Profit by Product</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Operating profit · column chart</div>', unsafe_allow_html=True)

        prod_profit = df.groupby("Product")["Operating Profit"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(6, 4))
        short_labels = [p.replace("Men's Street Footwear","M.Street FW")
                          .replace("Men's Athletic Footwear","M.Athletic FW")
                          .replace("Women's Street Footwear","W.Street FW")
                          .replace("Women's Athletic Footwear","W.Athletic FW")
                          .replace("Men's Apparel","M.Apparel")
                          .replace("Women's Apparel","W.Apparel")
                        for p in prod_profit.index]
        x = np.arange(len(prod_profit))
        bars = ax.bar(x, prod_profit.values, color=PALETTE[:len(prod_profit)],
                      width=0.62, edgecolor="none")
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, h * 1.01,
                    f"${h/1e6:.0f}M", ha='center', fontsize=7.5, color=TEXT_COLOR)
        ax.set_xticks(x)
        ax.set_xticklabels(short_labels, rotation=25, ha='right', fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, ylabel="Operating Profit (USD)")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col4:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Price vs Units Sold</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Scatter · color = Product category</div>', unsafe_allow_html=True)

        sample = df.sample(min(1500, len(df)), random_state=42)
        products_list = sorted(sample["Product"].unique())
        prod_colors   = {p: PALETTE[i % len(PALETTE)] for i, p in enumerate(products_list)}

        fig, ax = plt.subplots(figsize=(6, 4))
        for prod in products_list:
            sub = sample[sample["Product"] == prod]
            ax.scatter(sub["Price per Unit"], sub["Units Sold"],
                       color=prod_colors[prod], alpha=0.55, s=18,
                       edgecolors="none", label=prod.replace("Men's","M.").replace("Women's","W."))

        # Trend line
        m, b = np.polyfit(sample["Price per Unit"], sample["Units Sold"], 1)
        xr = np.linspace(sample["Price per Unit"].min(), sample["Price per Unit"].max(), 100)
        ax.plot(xr, m * xr + b, color=ACCENT, linewidth=1.5, linestyle="--", alpha=0.6)

        ax.legend(fontsize=6.5, framealpha=0, labelcolor=TEXT_COLOR, ncol=2, loc='upper right')
        fig_style(ax, xlabel="Price per Unit ($)", ylabel="Units Sold", grid_axis="both")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── ROW 3: Sales vs Profit scatter ────────────────────────────────────────
    col5, col6 = st.columns(2)

    with col5:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Sales vs Operating Profit</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Scatter · color = Retailer · size = Units Sold</div>', unsafe_allow_html=True)

        sample2 = df.sample(min(1500, len(df)), random_state=7)
        retailers_list = sorted(sample2["Retailer"].unique())
        ret_colors     = {r: PALETTE[i % len(PALETTE)] for i, r in enumerate(retailers_list)}

        fig, ax = plt.subplots(figsize=(6, 4))
        for ret in retailers_list:
            sub = sample2[sample2["Retailer"] == ret]
            sizes = (sub["Units Sold"] / sub["Units Sold"].max() * 60).clip(4, 80)
            ax.scatter(sub["Total Sales"], sub["Operating Profit"],
                       color=ret_colors[ret], alpha=0.55, s=sizes,
                       edgecolors="none", label=ret)

        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e3:.0f}K"))
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e3:.0f}K"))
        ax.legend(fontsize=7, framealpha=0, labelcolor=TEXT_COLOR, ncol=2)
        fig_style(ax, xlabel="Total Sales", ylabel="Operating Profit", grid_axis="both")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    with col6:
        st.markdown('<div class="chart-card">', unsafe_allow_html=True)
        st.markdown('<div class="chart-title">Retailer Profitability Ranking</div>', unsafe_allow_html=True)
        st.markdown('<div class="chart-subtitle">Total operating profit per retailer · column chart</div>', unsafe_allow_html=True)

        ret_profit = df.groupby("Retailer")["Operating Profit"].sum().sort_values(ascending=False)

        fig, ax = plt.subplots(figsize=(6, 4))
        x = np.arange(len(ret_profit))
        bars = ax.bar(x, ret_profit.values,
                      color=[PALETTE[i % len(PALETTE)] for i in range(len(ret_profit))],
                      width=0.62, edgecolor="none")
        for bar in bars:
            h = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, h * 1.01,
                    f"${h/1e6:.0f}M", ha='center', fontsize=8, color=TEXT_COLOR)
        ax.set_xticks(x)
        ax.set_xticklabels(ret_profit.index, rotation=20, ha='right', fontsize=9)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v/1e6:.0f}M"))
        fig_style(ax, ylabel="Operating Profit (USD)")
        plt.tight_layout()
        st.pyplot(fig)
        plt.close()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── KEY FINDINGS PAGE 2 ────────────────────────────────────────────────────
    st.markdown("<div class='section-divider'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div style='font-family:"Barlow Condensed",sans-serif;font-size:18px;font-weight:700;
                text-transform:uppercase;letter-spacing:2px;color:#7a9fff;margin-bottom:14px;'>
        🔍 Key Findings — Regional & Retail
    </div>
    """, unsafe_allow_html=True)

    top_city      = df.groupby("City")["Total Sales"].sum().idxmax()
    top_city_v    = df.groupby("City")["Total Sales"].sum().max()
    top_state_p   = df.groupby("State")["Total Sales"].sum().idxmax()
    low_price_ret = df.groupby("Retailer")["Price per Unit"].mean().idxmin()
    hi_price_ret  = df.groupby("Retailer")["Price per Unit"].mean().idxmax()
    corr_val      = df["Total Sales"].corr(df["Operating Profit"])
    best_method   = df.groupby("Sales Method")["Operating Profit"].sum().idxmax()

    findings2 = [
        ("Geographic Hotspot", f"<b>{top_city}</b> is the top city by sales (${top_city_v/1e6:.1f}M), making it Adidas's most critical urban market in the U.S."),
        ("State Leader", f"<b>{top_state_p}</b> is the highest-selling state, likely due to a dense retail network and high consumer spending capacity."),
        ("Profit-Sales Correlation", f"Sales and operating profit show a <b>{corr_val:.2f} correlation</b>, confirming that volume growth directly and consistently drives profitability."),
        ("Most Profitable Channel", f"<b>{best_method}</b> generates the highest operating profit among all sales channels, suggesting strong margins and execution."),
        ("Pricing Range", f"<b>{hi_price_ret}</b> has the highest average unit price, while <b>{low_price_ret}</b> operates at the lowest — indicating distinct market positioning strategies."),
        ("Retail Profitability", f"<b>{top_retailer2}</b> is the most profitable retail partner with ${top_ret_profit/1e6:.1f}M in operating profit, making it the most strategically valuable relationship for Adidas."),
    ]

    c_a2, c_b2 = st.columns(2)
    for i, (title, body) in enumerate(findings2):
        target = c_a2 if i % 2 == 0 else c_b2
        with target:
            st.markdown(f"""
            <div class="insight-box">
                <div class="insight-title">{title}</div>
                {body}
            </div>""", unsafe_allow_html=True)
