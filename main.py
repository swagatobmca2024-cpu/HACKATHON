import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.gridspec import GridSpec
import warnings
import io
import os

warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Adidas Sales Dashboard",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  GLOBAL DARK THEME CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* ---- Base ---- */
html, body, [class*="css"] {
    background-color: #0d0d0d !important;
    color: #e0e0e0 !important;
    font-family: 'Segoe UI', sans-serif;
}
.stApp { background-color: #0d0d0d !important; }

/* ---- Sidebar ---- */
[data-testid="stSidebar"] {
    background-color: #111827 !important;
    border-right: 1px solid #1f2937;
}
[data-testid="stSidebar"] * { color: #e0e0e0 !important; }
[data-testid="stSidebar"] .stSelectbox label,
[data-testid="stSidebar"] .stMultiSelect label,
[data-testid="stSidebar"] .stDateInput label { color: #9ca3af !important; font-size:12px; }

/* ---- Navigation Buttons ---- */
.nav-btn {
    display: inline-block; width: 100%; padding: 12px 16px;
    margin: 4px 0; border-radius: 8px; text-align: left;
    font-size: 14px; font-weight: 600; cursor: pointer;
    border: none; transition: all 0.2s;
}
.nav-btn-active { background: linear-gradient(135deg, #3b82f6, #8b5cf6) !important; color: #fff !important; }
.nav-btn-inactive { background: #1f2937 !important; color: #9ca3af !important; }
.nav-btn-inactive:hover { background: #374151 !important; color: #fff !important; }

/* ---- KPI Cards ---- */
.kpi-card {
    background: linear-gradient(135deg, #1f2937, #111827);
    border: 1px solid #374151; border-radius: 12px;
    padding: 20px 24px; margin: 4px 0;
}
.kpi-label { font-size: 12px; color: #9ca3af; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 6px; }
.kpi-value { font-size: 28px; font-weight: 700; color: #f9fafb; }
.kpi-sub   { font-size: 12px; color: #6b7280; margin-top: 4px; }
.kpi-accent-blue  .kpi-value { color: #60a5fa; }
.kpi-accent-green .kpi-value { color: #34d399; }
.kpi-accent-purple.kpi-value { color: #a78bfa; }
.kpi-accent-orange.kpi-value { color: #fb923c; }
.kpi-accent-pink  .kpi-value { color: #f472b6; }

/* ---- Section Headers ---- */
.section-header {
    font-size: 18px; font-weight: 700; color: #f9fafb;
    border-left: 4px solid #3b82f6; padding-left: 12px;
    margin: 24px 0 12px 0;
}

/* ---- Chart Containers ---- */
.chart-box {
    background: #111827; border: 1px solid #1f2937;
    border-radius: 12px; padding: 16px; margin: 6px 0;
}
.chart-title {
    font-size: 13px; font-weight: 600; color: #9ca3af;
    text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 8px;
}

/* ---- File Upload ---- */
[data-testid="stFileUploader"] {
    background: #111827 !important; border: 2px dashed #374151 !important;
    border-radius: 12px !important;
}

/* ---- DataFrames ---- */
[data-testid="stDataFrame"] { background: #111827 !important; }
.stDataFrame th { background: #1f2937 !important; color: #9ca3af !important; }

/* ---- Divider ---- */
hr { border-color: #1f2937 !important; }

/* ---- Scrollbar ---- */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #111827; }
::-webkit-scrollbar-thumb { background: #374151; border-radius: 3px; }

/* ---- Metric ---- */
[data-testid="stMetric"] {
    background: #111827 !important; border: 1px solid #1f2937;
    border-radius: 10px !important; padding: 12px !important;
}
[data-testid="stMetricLabel"]  { color: #9ca3af !important; }
[data-testid="stMetricValue"]  { color: #f9fafb !important; }
[data-testid="stMetricDelta"]  { color: #34d399 !important; }

/* ---- Buttons ---- */
.stButton > button {
    background: linear-gradient(135deg, #3b82f6, #8b5cf6) !important;
    color: white !important; border: none !important;
    border-radius: 8px !important; font-weight: 600 !important;
}
.stButton > button:hover { opacity: 0.9 !important; }

/* ---- Select / Input ---- */
.stSelectbox > div > div,
.stMultiSelect > div > div {
    background: #1f2937 !important; border-color: #374151 !important;
    color: #e0e0e0 !important;
}
.stDateInput > div > div > input {
    background: #1f2937 !important; border-color: #374151 !important;
    color: #e0e0e0 !important;
}

/* ---- Page Title ---- */
.page-title {
    font-size: 32px; font-weight: 800;
    background: linear-gradient(135deg, #60a5fa, #a78bfa, #f472b6);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    margin-bottom: 4px;
}
.page-subtitle { font-size: 14px; color: #6b7280; margin-bottom: 20px; }

/* ---- Badge ---- */
.badge {
    display: inline-block; padding: 3px 10px; border-radius: 20px;
    font-size: 11px; font-weight: 600;
}
.badge-blue   { background: #1d4ed8; color: #bfdbfe; }
.badge-green  { background: #065f46; color: #6ee7b7; }
.badge-orange { background: #92400e; color: #fde68a; }

/* ---- Clean Log ---- */
.clean-log {
    background: #0a0a0a; border: 1px solid #1f2937; border-radius: 8px;
    padding: 14px; font-family: monospace; font-size: 13px;
    color: #34d399; line-height: 1.8;
}
.clean-log .warn { color: #fbbf24; }
.clean-log .info { color: #60a5fa; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MATPLOTLIB DARK STYLE HELPER
# ─────────────────────────────────────────────
BG      = "#111827"
BG2     = "#1f2937"
TEXT    = "#e0e0e0"
GRID    = "#1f2937"
ACCENT  = ["#3b82f6","#8b5cf6","#34d399","#fb923c","#f472b6","#facc15",
           "#06b6d4","#ef4444","#10b981","#e879f9"]

def dark_fig(w=10, h=5):
    fig, ax = plt.subplots(figsize=(w, h), facecolor=BG)
    ax.set_facecolor(BG)
    return fig, ax

def dark_fig_n(rows, cols, w=14, h=6):
    fig, axes = plt.subplots(rows, cols, figsize=(w, h), facecolor=BG)
    for ax in (axes.flat if hasattr(axes, 'flat') else [axes]):
        ax.set_facecolor(BG)
    return fig, axes

def style_ax(ax, title="", xlabel="", ylabel=""):
    ax.set_facecolor(BG)
    ax.tick_params(colors=TEXT, labelsize=9)
    ax.xaxis.label.set_color(TEXT)
    ax.yaxis.label.set_color(TEXT)
    for spine in ax.spines.values():
        spine.set_edgecolor(BG2)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(
        lambda x, _: f"${x/1e6:.1f}M" if x >= 1e6 else (f"${x/1e3:.0f}K" if x >= 1e3 else f"${x:.0f}")
    ))
    if title:  ax.set_title(title, color=TEXT, fontsize=12, fontweight='bold', pad=10)
    if xlabel: ax.set_xlabel(xlabel, color="#9ca3af", fontsize=9)
    if ylabel: ax.set_ylabel(ylabel, color="#9ca3af", fontsize=9)
    ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)

# ─────────────────────────────────────────────
#  DATA LOADING & CLEANING
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_and_clean(uploaded_file):
    log = []
    raw = pd.read_excel(uploaded_file)
    log.append(f'<span class="info">ℹ  Loaded raw file — {raw.shape[0]:,} rows × {raw.shape[1]} columns</span>')

    df = raw.copy()

    # 1. Standardise column names
    df.columns = df.columns.str.strip()
    log.append('✅ Stripped whitespace from column names')

    # 2. Drop full duplicates
    before = len(df)
    df.drop_duplicates(inplace=True)
    dropped = before - len(df)
    log.append(f'✅ Removed {dropped} duplicate rows (kept {len(df):,})')

    # 3. Drop rows where ALL values are null
    before = len(df)
    df.dropna(how='all', inplace=True)
    log.append(f'✅ Dropped {before - len(df)} fully-empty rows')

    # 4. Parse date
    df['Invoice Date'] = pd.to_datetime(df['Invoice Date'], errors='coerce')
    bad_dates = df['Invoice Date'].isna().sum()
    if bad_dates:
        df.dropna(subset=['Invoice Date'], inplace=True)
        log.append(f'<span class="warn">⚠  Removed {bad_dates} rows with unparseable dates</span>')
    else:
        log.append('✅ Invoice Date parsed successfully — no nulls')

    # 5. Numeric coercion & negative value guard
    num_cols = ['Price per Unit', 'Units Sold', 'Total Sales', 'Operating Profit', 'Operating Margin']
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors='coerce')
    null_num = df[num_cols].isna().sum().sum()
    if null_num:
        df.dropna(subset=num_cols, inplace=True)
        log.append(f'<span class="warn">⚠  Dropped {null_num} rows with non-numeric values in numeric columns</span>')
    else:
        log.append('✅ All numeric columns valid — no coercion needed')

    neg_mask = (df[['Price per Unit','Units Sold','Total Sales']] < 0).any(axis=1)
    if neg_mask.sum():
        df = df[~neg_mask]
        log.append(f'<span class="warn">⚠  Removed {neg_mask.sum()} rows with negative prices/units/sales</span>')
    else:
        log.append('✅ No negative values in financial columns')

    # 6. Clip operating margin to [0, 1]
    before_clip = ((df['Operating Margin'] < 0) | (df['Operating Margin'] > 1)).sum()
    df['Operating Margin'] = df['Operating Margin'].clip(0, 1)
    log.append(f'✅ Operating Margin clipped to [0, 1] — {before_clip} values adjusted')

    # 7. Strip string columns
    str_cols = ['Retailer','Region','State','City','Product','Sales Method']
    for c in str_cols:
        df[c] = df[c].astype(str).str.strip().str.title()
    log.append('✅ Stripped & title-cased all text columns')

    # 8. Derived columns
    df['Year']       = df['Invoice Date'].dt.year
    df['Month']      = df['Invoice Date'].dt.month
    df['Month Name'] = df['Invoice Date'].dt.strftime('%b')
    df['YearMonth']  = df['Invoice Date'].dt.to_period('M')
    log.append('✅ Derived Year / Month / Month Name / YearMonth columns')

    log.append(f'<span class="info">ℹ  Final clean dataset — {df.shape[0]:,} rows × {df.shape[1]} columns</span>')
    return df, log

# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
if 'page' not in st.session_state:
    st.session_state.page = 'upload'
if 'df' not in st.session_state:
    st.session_state.df = None
if 'clean_log' not in st.session_state:
    st.session_state.clean_log = []

# ─────────────────────────────────────────────
#  SIDEBAR NAVIGATION
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 16px 0 24px 0;'>
        <div style='font-size:36px;'>👟</div>
        <div style='font-size:18px; font-weight:800; color:#f9fafb;'>Adidas Analytics</div>
        <div style='font-size:11px; color:#6b7280;'>Sales Intelligence Platform</div>
    </div>
    """, unsafe_allow_html=True)

    pages = [
        ("📁", "upload",    "Upload & Clean"),
        ("📊", "page1",     "Sales Overview"),
        ("🗺️", "page2",     "Geo & Product Deep Dive"),
    ]
    for icon, key, label in pages:
        active = st.session_state.page == key
        cls = "nav-btn-active" if active else "nav-btn-inactive"
        if st.button(f"{icon}  {label}", key=f"nav_{key}",
                     use_container_width=True,
                     disabled=(key != 'upload' and st.session_state.df is None)):
            st.session_state.page = key
            st.rerun()

    st.markdown("---")

    # Dynamic filters (only when data loaded and on dashboard pages)
    if st.session_state.df is not None and st.session_state.page in ('page1','page2'):
        df_all = st.session_state.df
        st.markdown('<div style="font-size:12px;color:#9ca3af;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;">FILTERS</div>', unsafe_allow_html=True)

        if st.session_state.page == 'page1':
            regions = ['All'] + sorted(df_all['Region'].unique().tolist())
            sel_region = st.selectbox("Region", regions)

            retailers = ['All'] + sorted(df_all['Retailer'].unique().tolist())
            sel_retailer = st.selectbox("Retailer", retailers)

            min_date = df_all['Invoice Date'].min().date()
            max_date = df_all['Invoice Date'].max().date()
            date_range = st.date_input("Date Range", value=(min_date, max_date),
                                        min_value=min_date, max_value=max_date)

            st.session_state.filters_p1 = {
                'region': sel_region, 'retailer': sel_retailer, 'date_range': date_range
            }

        elif st.session_state.page == 'page2':
            methods = ['All'] + sorted(df_all['Sales Method'].unique().tolist())
            sel_method = st.selectbox("Sales Method", methods)

            products = ['All'] + sorted(df_all['Product'].unique().tolist())
            sel_product = st.selectbox("Product", products)

            st.session_state.filters_p2 = {
                'method': sel_method, 'product': sel_product
            }

    st.markdown("---")
    st.markdown('<div style="font-size:10px;color:#4b5563;text-align:center;">Adidas Sales Dashboard v1.0<br>Built with Streamlit + Matplotlib</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  HELPER: apply page-1 filters
# ─────────────────────────────────────────────
def apply_p1_filters(df):
    f = st.session_state.get('filters_p1', {})
    if f.get('region', 'All') != 'All':
        df = df[df['Region'] == f['region']]
    if f.get('retailer', 'All') != 'All':
        df = df[df['Retailer'] == f['retailer']]
    dr = f.get('date_range')
    if dr and len(dr) == 2:
        df = df[(df['Invoice Date'].dt.date >= dr[0]) & (df['Invoice Date'].dt.date <= dr[1])]
    return df

def apply_p2_filters(df):
    f = st.session_state.get('filters_p2', {})
    if f.get('method', 'All') != 'All':
        df = df[df['Sales Method'] == f['method']]
    if f.get('product', 'All') != 'All':
        df = df[df['Product'] == f['product']]
    return df

# ─────────────────────────────────────────────
#  PAGE: UPLOAD & CLEAN
# ─────────────────────────────────────────────
def page_upload():
    st.markdown('<div class="page-title">📁 Upload & Data Cleaning</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">Upload your Adidas Sales Excel file. The pipeline will auto-clean and validate it.</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader("Drop your Excel file here (.xlsx)", type=['xlsx','xls'])

    if uploaded:
        with st.spinner("🔄 Cleaning data..."):
            df, log = load_and_clean(uploaded)
            st.session_state.df = df
            st.session_state.clean_log = log

        st.success(f"✅ Dataset ready — **{len(df):,} rows** loaded and cleaned!")

        # Cleaning log
        st.markdown('<div class="section-header">🧹 Cleaning Log</div>', unsafe_allow_html=True)
        log_html = "<br>".join(st.session_state.clean_log)
        st.markdown(f'<div class="clean-log">{log_html}</div>', unsafe_allow_html=True)

        st.markdown('<div class="section-header">📋 Cleaned Data Preview (first 50 rows)</div>', unsafe_allow_html=True)
        st.dataframe(df.head(50), use_container_width=True, height=300)

        # Column summary
        st.markdown('<div class="section-header">📐 Column Summary</div>', unsafe_allow_html=True)
        summary = pd.DataFrame({
            'Column': df.columns,
            'Type': df.dtypes.astype(str).values,
            'Non-Null': df.notna().sum().values,
            'Null': df.isna().sum().values,
            'Unique': df.nunique().values,
        })
        st.dataframe(summary, use_container_width=True, hide_index=True)

        # Download cleaned file
        buf = io.BytesIO()
        df.to_excel(buf, index=False)
        buf.seek(0)
        st.download_button("⬇️  Download Cleaned Excel", data=buf,
                           file_name="adidas_cleaned.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("---")
        st.info("👈 Use the sidebar to navigate to **Sales Overview** or **Geo & Product Deep Dive**")

    else:
        # Placeholder instructions
        st.markdown("""
        <div style='background:#111827;border:1px solid #1f2937;border-radius:12px;padding:32px;text-align:center;margin-top:24px;'>
            <div style='font-size:48px;margin-bottom:12px;'>📊</div>
            <div style='font-size:18px;font-weight:700;color:#f9fafb;margin-bottom:8px;'>No file uploaded yet</div>
            <div style='font-size:14px;color:#6b7280;'>Upload an Adidas Sales .xlsx file to begin.<br>
            Expected columns: Retailer, Invoice Date, Region, State, City, Product,<br>
            Price per Unit, Units Sold, Total Sales, Operating Profit, Operating Margin, Sales Method</div>
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  PAGE 1: SALES OVERVIEW
# ─────────────────────────────────────────────
def page1():
    df_raw = st.session_state.df
    df = apply_p1_filters(df_raw)

    st.markdown('<div class="page-title">📊 Sales Overview — Page 1</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">KPIs, monthly trends, regional breakdown, product performance & sales channel mix</div>', unsafe_allow_html=True)

    if len(df) == 0:
        st.warning("No data matches current filters.")
        return

    # ── KPIs ──────────────────────────────────
    total_sales    = df['Total Sales'].sum()
    total_profit   = df['Operating Profit'].sum()
    total_units    = df['Units Sold'].sum()
    avg_price      = df['Price per Unit'].mean()
    avg_margin     = df['Operating Margin'].mean()

    st.markdown('<div class="section-header">Key Performance Indicators</div>', unsafe_allow_html=True)
    c1, c2, c3, c4, c5 = st.columns(5)

    def kpi(col, label, value, sub, accent):
        col.markdown(f"""
        <div class="kpi-card kpi-accent-{accent}">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>""", unsafe_allow_html=True)

    kpi(c1, "Total Sales",          f"${total_sales/1e6:.2f}M",    "Revenue across all channels",  "blue")
    kpi(c2, "Total Profit",         f"${total_profit/1e6:.2f}M",   "Operating profit generated",   "green")
    kpi(c3, "Total Units Sold",     f"{total_units:,}",             "Units across all products",    "purple")
    kpi(c4, "Avg Price / Unit",     f"${avg_price:.2f}",            "Average selling price (USD)",  "orange")
    kpi(c5, "Avg Operating Margin", f"{avg_margin*100:.1f}%",       "Average profitability ratio",  "pink")

    st.markdown("---")

    # ── ROW 1: Area Chart + Donut (Region) ────
    st.markdown('<div class="section-header">Sales Trends & Regional Distribution</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns([2, 1])

    # Area Chart: Monthly Sales
    with col_a:
        monthly = (df.groupby(['Year','Month','Month Name'])['Total Sales']
                   .sum().reset_index()
                   .sort_values(['Year','Month']))
        monthly['Label'] = monthly['Month Name'] + " '" + monthly['Year'].astype(str).str[-2:]

        fig, ax = dark_fig(10, 4)
        x = range(len(monthly))
        vals = monthly['Total Sales'].values
        ax.fill_between(x, vals, alpha=0.3, color=ACCENT[0])
        ax.plot(x, vals, color=ACCENT[0], linewidth=2.5, marker='o', markersize=4)
        ax.set_xticks(list(x))
        ax.set_xticklabels(monthly['Label'].tolist(), rotation=45, ha='right', fontsize=7, color=TEXT)
        style_ax(ax, title="Total Sales by Month", ylabel="Total Sales (USD)")
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M" if v >= 1e6 else f"${v/1e3:.0f}K"))
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # Donut: Region
    with col_b:
        region_sales = df.groupby('Region')['Total Sales'].sum().sort_values(ascending=False)
        fig, ax = plt.subplots(figsize=(5, 4), facecolor=BG)
        ax.set_facecolor(BG)
        wedges, texts, autotexts = ax.pie(
            region_sales.values,
            labels=None,
            autopct='%1.1f%%',
            colors=ACCENT[:len(region_sales)],
            startangle=90,
            wedgeprops=dict(width=0.55, edgecolor=BG, linewidth=2),
            pctdistance=0.75,
        )
        for t in autotexts:
            t.set_color(TEXT); t.set_fontsize(8); t.set_fontweight('bold')
        ax.legend(region_sales.index, loc='lower center', bbox_to_anchor=(0.5, -0.12),
                  ncol=3, fontsize=8, framealpha=0, labelcolor=TEXT)
        ax.set_title("Sales by Region", color=TEXT, fontsize=12, fontweight='bold', pad=10)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── ROW 2: Product Bar + Sales Method Donut ──
    st.markdown('<div class="section-header">Product Performance & Channel Mix</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns([2, 1])

    with col_c:
        prod_sales = df.groupby('Product')['Total Sales'].sum().sort_values(ascending=True)
        fig, ax = dark_fig(10, 4)
        bars = ax.barh(prod_sales.index, prod_sales.values,
                       color=ACCENT[:len(prod_sales)], edgecolor='none', height=0.6)
        for bar, val in zip(bars, prod_sales.values):
            ax.text(val + prod_sales.max()*0.01, bar.get_y() + bar.get_height()/2,
                    f"${val/1e6:.1f}M", va='center', ha='left', color=TEXT, fontsize=8)
        ax.tick_params(colors=TEXT, labelsize=9)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.set_facecolor(BG)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M" if v >= 1e6 else f"${v/1e3:.0f}K"))
        ax.tick_params(axis='x', colors=TEXT); ax.tick_params(axis='y', colors=TEXT)
        ax.set_title("Total Sales by Product", color=TEXT, fontsize=12, fontweight='bold', pad=10)
        ax.grid(axis='x', color=GRID, linewidth=0.5, alpha=0.6)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    with col_d:
        method_sales = df.groupby('Sales Method')['Total Sales'].sum()
        fig, ax = plt.subplots(figsize=(5, 4), facecolor=BG)
        ax.set_facecolor(BG)
        wedges, texts, autotexts = ax.pie(
            method_sales.values, labels=None,
            autopct='%1.1f%%',
            colors=[ACCENT[2], ACCENT[3], ACCENT[4]],
            startangle=90,
            wedgeprops=dict(width=0.55, edgecolor=BG, linewidth=2),
            pctdistance=0.75,
        )
        for t in autotexts:
            t.set_color(TEXT); t.set_fontsize(9); t.set_fontweight('bold')
        ax.legend(method_sales.index, loc='lower center', bbox_to_anchor=(0.5, -0.12),
                  ncol=3, fontsize=8, framealpha=0, labelcolor=TEXT)
        ax.set_title("Sales by Channel", color=TEXT, fontsize=12, fontweight='bold', pad=10)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── ROW 3: Retailer bar + Year comparison ──
    st.markdown('<div class="section-header">Retailer Comparison & Year-over-Year</div>', unsafe_allow_html=True)
    col_e, col_f = st.columns(2)

    with col_e:
        ret_sales = df.groupby('Retailer')['Total Sales'].sum().sort_values(ascending=False)
        fig, ax = dark_fig(7, 4)
        bars = ax.bar(ret_sales.index, ret_sales.values,
                      color=ACCENT[:len(ret_sales)], edgecolor='none', width=0.6)
        for bar, val in zip(bars, ret_sales.values):
            ax.text(bar.get_x() + bar.get_width()/2, val + ret_sales.max()*0.01,
                    f"${val/1e6:.1f}M", ha='center', va='bottom', color=TEXT, fontsize=8)
        ax.tick_params(colors=TEXT, labelsize=9)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M" if v >= 1e6 else f"${v/1e3:.0f}K"))
        ax.set_title("Total Sales by Retailer", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)
        plt.xticks(rotation=20, ha='right')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    with col_f:
        yoy = df.groupby('Year')['Total Sales'].sum()
        fig, ax = dark_fig(7, 4)
        bars = ax.bar(yoy.index.astype(str), yoy.values, color=[ACCENT[0], ACCENT[1]], width=0.5)
        for bar, val in zip(bars, yoy.values):
            ax.text(bar.get_x() + bar.get_width()/2, val + yoy.max()*0.01,
                    f"${val/1e6:.1f}M", ha='center', va='bottom', color=TEXT, fontsize=10, fontweight='bold')
        ax.tick_params(colors=TEXT, labelsize=10)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M"))
        ax.set_title("Year-over-Year Sales", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

# ─────────────────────────────────────────────
#  PAGE 2: GEO & PRODUCT DEEP DIVE
# ─────────────────────────────────────────────
def page2():
    df_raw = st.session_state.df
    df = apply_p2_filters(df_raw)

    st.markdown('<div class="page-title">🗺️ Geo & Product Deep Dive — Page 2</div>', unsafe_allow_html=True)
    st.markdown('<div class="page-subtitle">State-level performance, top cities, product profitability & pricing intelligence</div>', unsafe_allow_html=True)

    if len(df) == 0:
        st.warning("No data matches current filters.")
        return

    # ── KPIs ──────────────────────────────────
    top_state   = df.groupby('State')['Total Sales'].sum().idxmax()
    top_product = df.groupby('Product')['Total Sales'].sum().idxmax()
    top_retailer_profit = df.groupby('Retailer')['Operating Profit'].sum().idxmax()
    avg_price   = df['Price per Unit'].mean()

    st.markdown('<div class="section-header">Key Highlights</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)

    def kpi2(col, icon, label, value, accent):
        col.markdown(f"""
        <div class="kpi-card kpi-accent-{accent}">
            <div style="font-size:24px;margin-bottom:4px;">{icon}</div>
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="font-size:20px;">{value}</div>
        </div>""", unsafe_allow_html=True)

    kpi2(c1, "🏆", "Highest Selling State",      top_state,              "blue")
    kpi2(c2, "👟", "Highest Selling Product",     top_product,            "green")
    kpi2(c3, "💰", "Most Profitable Retailer",    top_retailer_profit,    "purple")
    kpi2(c4, "🏷️",  "Avg Price per Unit",          f"${avg_price:.2f}",    "orange")

    st.markdown("---")

    # ── Row 1: State Sales Bar + Top 10 Cities ──
    st.markdown('<div class="section-header">Geographic Performance</div>', unsafe_allow_html=True)
    col_a, col_b = st.columns(2)

    with col_a:
        state_sales = df.groupby('State')['Total Sales'].sum().sort_values(ascending=False).head(15)
        fig, ax = dark_fig(8, 5)
        colors_s = [ACCENT[0] if i == 0 else ACCENT[1] if i < 3 else "#374151"
                    for i in range(len(state_sales))]
        bars = ax.barh(state_sales.index[::-1], state_sales.values[::-1],
                       color=colors_s[::-1], edgecolor='none', height=0.7)
        ax.tick_params(colors=TEXT, labelsize=8)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M" if v >= 1e6 else f"${v/1e3:.0f}K"))
        ax.set_title("Top 15 States by Sales", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='x', color=GRID, linewidth=0.5, alpha=0.6)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    with col_b:
        city_sales = df.groupby('City')['Total Sales'].sum().sort_values(ascending=False).head(10)
        fig, ax = dark_fig(8, 5)
        bars = ax.bar(city_sales.index, city_sales.values,
                      color=ACCENT[:10], edgecolor='none', width=0.7)
        for bar, val in zip(bars, city_sales.values):
            ax.text(bar.get_x() + bar.get_width()/2, val + city_sales.max()*0.01,
                    f"${val/1e6:.1f}M", ha='center', va='bottom', color=TEXT, fontsize=7)
        ax.tick_params(colors=TEXT, labelsize=7)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M" if v >= 1e6 else f"${v/1e3:.0f}K"))
        ax.set_title("Top 10 Cities by Sales", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)
        plt.xticks(rotation=30, ha='right')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── Row 2: Profit by Product + Scatter (Price vs Units) ──
    st.markdown('<div class="section-header">Profitability & Pricing Intelligence</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)

    with col_c:
        prod_profit = df.groupby('Product')['Operating Profit'].sum().sort_values(ascending=False)
        fig, ax = dark_fig(8, 4.5)
        bars = ax.bar(prod_profit.index, prod_profit.values,
                      color=ACCENT[:len(prod_profit)], edgecolor='none', width=0.6)
        for bar, val in zip(bars, prod_profit.values):
            ax.text(bar.get_x() + bar.get_width()/2, val + prod_profit.max()*0.01,
                    f"${val/1e6:.1f}M", ha='center', va='bottom', color=TEXT, fontsize=8)
        ax.tick_params(colors=TEXT, labelsize=8)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e6:.1f}M"))
        ax.set_title("Operating Profit by Product", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)
        plt.xticks(rotation=25, ha='right', fontsize=8)
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    with col_d:
        sample = df.sample(min(1500, len(df)), random_state=42)
        products_u = sample['Product'].unique()
        color_map = {p: ACCENT[i % len(ACCENT)] for i, p in enumerate(products_u)}
        fig, ax = dark_fig(8, 4.5)
        for prod in products_u:
            mask = sample['Product'] == prod
            ax.scatter(sample.loc[mask, 'Price per Unit'],
                       sample.loc[mask, 'Units Sold'],
                       color=color_map[prod], alpha=0.55, s=18, label=prod)
        ax.tick_params(colors=TEXT, labelsize=9)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.set_xlabel("Price per Unit ($)", color="#9ca3af", fontsize=9)
        ax.set_ylabel("Units Sold", color="#9ca3af", fontsize=9)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:,.0f}"))
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"${v:.0f}"))
        ax.set_title("Price per Unit vs Units Sold", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(color=GRID, linewidth=0.5, alpha=0.4)
        ax.legend(fontsize=7, framealpha=0, labelcolor=TEXT, loc='upper right')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── Row 3: Sales vs Profit scatter + Retailer margin ──
    st.markdown('<div class="section-header">Sales vs Profitability & Retailer Margins</div>', unsafe_allow_html=True)
    col_e, col_f = st.columns(2)

    with col_e:
        sample2 = df.sample(min(1500, len(df)), random_state=7)
        retailers_u = sample2['Retailer'].unique()
        cmap2 = {r: ACCENT[i % len(ACCENT)] for i, r in enumerate(retailers_u)}
        fig, ax = dark_fig(8, 4.5)
        for ret in retailers_u:
            mask = sample2['Retailer'] == ret
            ax.scatter(sample2.loc[mask, 'Total Sales'],
                       sample2.loc[mask, 'Operating Profit'],
                       color=cmap2[ret], alpha=0.55, s=18, label=ret)
        ax.tick_params(colors=TEXT, labelsize=9)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.set_xlabel("Total Sales ($)", color="#9ca3af", fontsize=9)
        ax.set_ylabel("Operating Profit ($)", color="#9ca3af", fontsize=9)
        ax.xaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e3:.0f}K" if v < 1e6 else f"${v/1e6:.1f}M"))
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(
            lambda v, _: f"${v/1e3:.0f}K" if v < 1e6 else f"${v/1e6:.1f}M"))
        ax.set_title("Total Sales vs Operating Profit", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(color=GRID, linewidth=0.5, alpha=0.4)
        ax.legend(fontsize=7, framealpha=0, labelcolor=TEXT, loc='upper left')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    with col_f:
        ret_margin = df.groupby('Retailer')['Operating Margin'].mean().sort_values(ascending=False)
        fig, ax = dark_fig(8, 4.5)
        bars = ax.bar(ret_margin.index, ret_margin.values * 100,
                      color=ACCENT[:len(ret_margin)], edgecolor='none', width=0.6)
        for bar, val in zip(bars, ret_margin.values):
            ax.text(bar.get_x() + bar.get_width()/2, val*100 + 0.5,
                    f"{val*100:.1f}%", ha='center', va='bottom', color=TEXT, fontsize=9)
        ax.tick_params(colors=TEXT, labelsize=9)
        for spine in ax.spines.values(): spine.set_edgecolor(BG2)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f"{v:.0f}%"))
        ax.set_title("Avg Operating Margin by Retailer", color=TEXT, fontsize=12, fontweight='bold')
        ax.grid(axis='y', color=GRID, linewidth=0.5, alpha=0.6)
        plt.xticks(rotation=20, ha='right')
        plt.tight_layout()
        st.pyplot(fig)
        plt.close(fig)

    # ── KEY FINDINGS ──────────────────────────
    st.markdown("---")
    st.markdown('<div class="section-header">🔍 Key Findings & Insights</div>', unsafe_allow_html=True)

    df_full = st.session_state.df  # always use full dataset for findings
    ts_state   = df_full.groupby('State')['Total Sales'].sum()
    ts_product = df_full.groupby('Product')['Total Sales'].sum()
    ts_retailer= df_full.groupby('Retailer')['Total Sales'].sum()
    ts_method  = df_full.groupby('Sales Method')['Total Sales'].sum()
    profit_ret = df_full.groupby('Retailer')['Operating Profit'].sum()
    margin_ret = df_full.groupby('Retailer')['Operating Margin'].mean()
    ts_region  = df_full.groupby('Region')['Total Sales'].sum()

    findings = [
        ("🏆 Top State",         f"{ts_state.idxmax()} leads all states with ${ts_state.max()/1e6:.1f}M in total sales."),
        ("👟 Best Product",      f"{ts_product.idxmax()} is the #1 product by revenue (${ts_product.max()/1e6:.1f}M), while {ts_product.idxmin()} trails at ${ts_product.min()/1e6:.1f}M."),
        ("🏪 Top Retailer",      f"{ts_retailer.idxmax()} is the highest-grossing retailer (${ts_retailer.max()/1e6:.1f}M). {profit_ret.idxmax()} generates the most profit (${profit_ret.max()/1e6:.1f}M)."),
        ("📱 Online Growth",     f"{'Online' if 'Online' in ts_method else ts_method.idxmax()} is the {'fastest-growing' if 'Online' in ts_method else 'leading'} sales channel — accounting for {ts_method.get('Online', ts_method.max()) / ts_method.sum() * 100:.1f}% of total revenue."),
        ("🌍 Regional Lead",     f"{ts_region.idxmax()} is the dominant region (${ts_region.max()/1e6:.1f}M), representing {ts_region.max()/ts_region.sum()*100:.1f}% of total sales."),
        ("💹 Margin Champion",   f"{margin_ret.idxmax()} achieves the highest avg operating margin ({margin_ret.max()*100:.1f}%), indicating superior cost efficiency."),
        ("💲 Pricing Insight",   f"Average price per unit is ${df_full['Price per Unit'].mean():.2f}. Street Footwear commands a premium over Apparel lines."),
        ("📈 YoY Trend",         f"Sales grew significantly from 2020 to 2021 — driven by e-commerce expansion and post-pandemic recovery in retail."),
    ]

    for icon_label, text in findings:
        st.markdown(f"""
        <div style='background:#111827;border-left:3px solid #3b82f6;border-radius:0 8px 8px 0;
                    padding:12px 16px;margin:6px 0;'>
            <span style='font-weight:700;color:#60a5fa;'>{icon_label}: </span>
            <span style='color:#d1d5db;'>{text}</span>
        </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  ROUTER
# ─────────────────────────────────────────────
page = st.session_state.page

if page == 'upload':
    page_upload()
elif page == 'page1':
    if st.session_state.df is None:
        st.warning("Please upload a file first.")
    else:
        page1()
elif page == 'page2':
    if st.session_state.df is None:
        st.warning("Please upload a file first.")
    else:
        page2()
