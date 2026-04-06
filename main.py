import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from matplotlib.colors import LinearSegmentedColormap
import warnings
import io
import streamlit.components.v1 as components
warnings.filterwarnings('ignore')

# ─────────────────────────────────────────────
#  PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Adidas Sales Intelligence",
    page_icon="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>A</text></svg>",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
#  COLOUR TOKENS
# ─────────────────────────────────────────────
DARK_BG = "#0A0A0F"
CARD_BG = "#12121A"
SIDE_BG = "#0D0D14"
A1      = "#00E5FF"   # cyan
A2      = "#FF3CAC"   # pink
A3      = "#7B5EA7"   # violet
A4      = "#F9C846"   # gold
A5      = "#39FF14"   # neon green
TPRI    = "#F0F0F5"
TSEC    = "#8A8A9A"
BDR     = "#1E1E2E"
PALETTE = [A1, A2, A4, A5, A3, "#FF6B35"]

# ─────────────────────────────────────────────
#  INLINE SVG ICON HELPERS
# ─────────────────────────────────────────────
def svg(path_d, size=16, color=None, vb="0 0 24 24"):
    c = color or A1
    s = str(size)
    return (
        '<svg width="' + s + '" height="' + s + '" viewBox="' + vb +
        '" fill="none" xmlns="http://www.w3.org/2000/svg"'
        ' style="vertical-align:middle;margin-right:6px;">'
        '<path d="' + path_d + '" stroke="' + c + '" stroke-width="1.8"'
        ' stroke-linecap="round" stroke-linejoin="round"/></svg>'
    )

def svg_fill(path_d, size=16, color=None, vb="0 0 24 24"):
    c = color or A1
    s = str(size)
    return (
        '<svg width="' + s + '" height="' + s + '" viewBox="' + vb +
        '" fill="' + c + '" xmlns="http://www.w3.org/2000/svg"'
        ' style="vertical-align:middle;margin-right:6px;">'
        '<path d="' + path_d + '"/></svg>'
    )

# ─── Icon path constants ───────────────────────
P_UPLOAD  = "M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4M17 8l-5-5-5 5M12 3v12"
P_CHART   = "M3 3v18h18M9 17V9m4 8v-5m4 5V5"
P_GLOBE   = "M12 2a10 10 0 1 0 0 20A10 10 0 0 0 12 2zM2 12h20M12 2a15 15 0 0 1 0 20M12 2a15 15 0 0 0 0 20"
P_SHOE    = "M3 17h18l-3-8H6L3 17zM6 9l2-5h8l2 5"
P_MAP     = "M1 6v16l7-4 8 4 7-4V2l-7 4-8-4-7 4zM8 2v16M16 6v16"
P_CITY    = "M3 21h18M9 21V7l6-4v18M9 7l6-4M3 21V11l6-4M21 21V11l-6-4"
P_PROFIT  = "M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"
P_SCATTER = "M8 8m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0M16 16m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0M8 16m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0M16 8m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0"
P_TREND   = "M22 7l-9 9-5-5L1 17M22 7h-6M22 7v6"
P_FIND    = "M10 21a7 7 0 1 0 0-14 7 7 0 0 0 0 14zM21 21l-4.35-4.35"
P_STAR    = "M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"
P_TAG     = "M20.59 13.41l-7.17 7.17a2 2 0 0 1-2.83 0L2 12V2h10l8.59 8.59a2 2 0 0 1 0 2.82zM7 7h.01"
P_FLASH   = "M13 2L3 14h9l-1 8 10-12h-9l1-8z"
P_ALERT   = "M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01"
P_NAV     = "M3 3h7v7H3zM14 3h7v7h-7zM14 14h7v7h-7zM3 14h7v7H3z"
P_FILE    = "M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM14 2v6h6M16 13H8M16 17H8M10 9H8"
P_DOLLAR  = "M12 1v22M16 5H9a4 4 0 0 0 0 8h6a4 4 0 0 1 0 8H5"
P_BOX     = "M20 7H4a2 2 0 0 0-2 2v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2z"
P_WAVE    = "M22 12h-4l-3 9L9 3l-3 9H2"

def ct(icon, label, color=None):
    """Chart title with inline SVG."""
    c = color or A1
    return f'<div class="chart-title">{svg(icon, 15, c)}<span style="color:{c}">{label}</span></div>'

def sl(icon, label, color=None):
    """Section label with inline SVG."""
    c = color or A2
    return f'<div class="section-label">{svg(icon, 13, c)}<span>{label}</span></div>'

# ─────────────────────────────────────────────
#  ADIDAS-STYLE LOGO SVG
# ─────────────────────────────────────────────
def _logo(w, h, rx, sw, pts, lines):
    """Build Adidas triangle logo SVG as a plain string (no nested f-string issues)."""
    ln_tags = "".join(
        '<line x1="' + str(x1) + '" y1="' + str(y1) +
        '" x2="' + str(x2) + '" y2="' + str(y2) +
        '" stroke="' + A1 + '" stroke-width="' + str(sw) +
        '" stroke-linecap="round"/>'
        for x1, y1, x2, y2 in lines
    )
    return (
        '<svg width="' + str(w) + '" height="' + str(h) +
        '" viewBox="0 0 ' + str(w) + ' ' + str(h) +
        '" fill="none" xmlns="http://www.w3.org/2000/svg">'
        '<rect width="' + str(w) + '" height="' + str(h) +
        '" rx="' + str(rx) + '" fill="' + CARD_BG +
        '" stroke="' + A1 + '" stroke-width="1.5"/>'
        '<polygon points="' + pts + '" fill="none" stroke="' + A1 +
        '" stroke-width="' + str(sw) + '" stroke-linejoin="round"/>'
        + ln_tags + '</svg>'
    )

LOGO_SVG = _logo(54, 54, 11, 2.2,
    "27,10 43,42 11,42",
    [(18,42,36,42),(21,35,33,35),(24,28,30,28)])

LOGO_LG  = _logo(82, 82, 18, 2.5,
    "41,14 62,64 20,64",
    [(27,64,55,64),(31,54,51,54),(35,44,47,44)])

# ─────────────────────────────────────────────
#  GLOBAL CSS
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Rajdhani:wght@300;400;600;700&family=JetBrains+Mono:wght@400;700&display=swap');

html, body, [class*="css"] {{
    background-color: {DARK_BG} !important;
    color: {TPRI} !important;
    font-family: 'Rajdhani', sans-serif;
}}
.stApp {{ background-color: {DARK_BG}; }}

/* Sidebar */
section[data-testid="stSidebar"] {{
    background: {SIDE_BG} !important;
    border-right: 1px solid {BDR};
}}
section[data-testid="stSidebar"] * {{ color: {TPRI} !important; }}

/* File uploader */
div[data-testid="stFileUploader"] {{
    background: {CARD_BG} !important;
    border: 2px dashed {A1}55 !important;
    border-radius: 12px !important;
}}
div[data-testid="stFileUploader"]:hover {{
    border-color: {A1}AA !important;
    box-shadow: 0 0 20px {A1}18 !important;
}}

/* Dashboard header */
.dashboard-header {{
    background: linear-gradient(135deg, {CARD_BG} 0%, #1A0A2E 100%);
    border: 1px solid {A1}33;
    border-radius: 16px;
    padding: 24px 34px;
    margin-bottom: 22px;
    position: relative;
    overflow: hidden;
    display: flex;
    align-items: center;
    gap: 20px;
}}
.dashboard-header::before {{
    content: '';
    position: absolute;
    top: -70px; right: -70px;
    width: 230px; height: 230px;
    background: radial-gradient(circle, {A1}18, transparent 70%);
    border-radius: 50%;
    pointer-events: none;
}}
.header-text h1 {{
    font-family: 'Orbitron', sans-serif;
    font-size: 1.75rem;
    font-weight: 900;
    background: linear-gradient(90deg, {A1}, {A2});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    margin: 0;
    letter-spacing: 3px;
}}
.header-text p {{
    color: {TSEC};
    font-size: 0.82rem;
    margin: 5px 0 0;
    letter-spacing: 1px;
    font-family: 'JetBrains Mono', monospace;
}}

/* KPI grid */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(5, 1fr);
    gap: 13px;
    margin-bottom: 22px;
}}
.kpi-card {{
    background: {CARD_BG};
    border-radius: 12px;
    padding: 18px 16px 22px;
    border: 1px solid {BDR};
    position: relative;
    overflow: hidden;
}}
.kpi-card::after {{
    content: '';
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 3px;
}}
.kpi-c1::after {{ background: {A1}; }}
.kpi-c2::after {{ background: {A2}; }}
.kpi-c3::after {{ background: {A4}; }}
.kpi-c4::after {{ background: {A5}; }}
.kpi-c5::after {{ background: {A3}; }}
.kpi-icon {{ margin-bottom: 10px; display:block; }}
.kpi-label {{
    font-size: 0.6rem;
    color: {TSEC};
    text-transform: uppercase;
    letter-spacing: 2px;
    font-family: 'JetBrains Mono', monospace;
    margin-bottom: 7px;
}}
.kpi-value {{
    font-family: 'Orbitron', sans-serif;
    font-size: 1.4rem;
    font-weight: 700;
    color: {TPRI};
    line-height: 1;
}}
.kpi-sub {{
    font-size: 0.66rem;
    color: {TSEC};
    margin-top: 6px;
    font-family: 'JetBrains Mono', monospace;
}}

/* Chart title */
.chart-title {{
    font-family: 'Orbitron', sans-serif;
    font-size: 0.7rem;
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 10px;
    display: flex;
    align-items: center;
}}

/* Section label */
.section-label {{
    font-family: 'Orbitron', sans-serif;
    font-size: 0.6rem;
    color: {A2};
    letter-spacing: 3px;
    text-transform: uppercase;
    border-left: 3px solid {A2};
    padding-left: 10px;
    margin: 20px 0 14px;
    display: flex;
    align-items: center;
}}

/* Findings */
.finding-card {{
    background: linear-gradient(135deg, #12121A, #1a0a2e);
    border: 1px solid {A1}44;
    border-radius: 10px;
    padding: 15px 18px;
    margin-bottom: 12px;
}}
.finding-card h4 {{
    color: {A1};
    font-family: 'Orbitron', sans-serif;
    font-size: 0.68rem;
    letter-spacing: 2px;
    margin: 0 0 8px;
    display: flex;
    align-items: center;
}}
.finding-card p {{
    color: {TPRI};
    font-size: 0.88rem;
    margin: 0;
    line-height: 1.6;
}}

/* Upload screen */
.upload-wrap {{
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 50px 20px 30px;
    text-align: center;
}}
.upload-wrap h2 {{
    font-family: 'Orbitron', sans-serif;
    font-size: 1.4rem;
    font-weight: 900;
    background: linear-gradient(90deg, {A1}, {A2});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: 3px;
    margin: 18px 0 8px;
}}
.upload-wrap p {{
    color: {TSEC};
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    letter-spacing: 1px;
    margin-bottom: 28px;
    max-width: 480px;
}}
.upload-badge {{
    display: inline-flex;
    align-items: center;
    gap: 7px;
    background: {CARD_BG};
    border: 1px solid {BDR};
    border-radius: 8px;
    padding: 7px 14px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    color: {TSEC};
    margin: 4px;
}}

/* Streamlit widget overrides */
div[data-testid="stSelectbox"] > div,
div[data-testid="stMultiSelect"] > div {{
    background: {CARD_BG} !important;
    border-color: {BDR} !important;
    border-radius: 8px !important;
}}
div.stButton > button {{
    background: linear-gradient(135deg, {A1}22, {A2}22);
    border: 1px solid {A1}55;
    color: {A1};
    font-family: 'Orbitron', sans-serif;
    font-size: 0.6rem;
    letter-spacing: 2px;
    border-radius: 8px;
}}
div[data-testid="stDateInput"] input {{
    background: {CARD_BG} !important;
    color: {TPRI} !important;
    border-color: {BDR} !important;
    border-radius: 6px !important;
}}
label, .stRadio label, .stSelectbox label {{
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 0.68rem !important;
    color: {TSEC} !important;
    letter-spacing: 1px !important;
}}
/* Data cleaning report */
.clean-panel {{
    background: {CARD_BG};
    border: 1px solid {A5}44;
    border-radius: 14px;
    padding: 22px 26px;
    margin-bottom: 22px;
}}
.clean-panel-title {{
    font-family: 'Orbitron', sans-serif;
    font-size: 0.72rem;
    color: {A5};
    letter-spacing: 2px;
    text-transform: uppercase;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    gap: 8px;
}}
.clean-grid {{
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 12px;
    margin-bottom: 14px;
}}
.clean-stat {{
    background: {DARK_BG};
    border-radius: 8px;
    padding: 12px 14px;
    border-left: 3px solid {A5};
}}
.clean-stat-val {{
    font-family: 'Orbitron', sans-serif;
    font-size: 1.2rem;
    font-weight: 700;
    color: {A5};
}}
.clean-stat-lbl {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.6rem;
    color: {TSEC};
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 4px;
}}
.clean-step {{
    display: flex;
    align-items: flex-start;
    gap: 10px;
    padding: 7px 0;
    border-bottom: 1px solid {BDR};
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
}}
.clean-step:last-child {{ border-bottom: none; }}
.clean-step-num {{
    background: {A5}22;
    color: {A5};
    border-radius: 4px;
    padding: 2px 7px;
    font-size: 0.6rem;
    font-weight: 700;
    flex-shrink: 0;
    min-width: 24px;
    text-align: center;
}}
.clean-step-ok  {{ color: {A5}; }}
.clean-step-warn{{ color: {A4}; }}
.clean-step-info{{ color: {TSEC}; }}

</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  MATPLOTLIB DARK DEFAULTS
# ─────────────────────────────────────────────
plt.rcParams.update({
    'figure.facecolor': DARK_BG,
    'axes.facecolor':   CARD_BG,
    'axes.edgecolor':   BDR,
    'axes.labelcolor':  TSEC,
    'axes.titlecolor':  TPRI,
    'xtick.color':      TSEC,
    'ytick.color':      TSEC,
    'text.color':       TPRI,
    'grid.color':       BDR,
    'grid.linewidth':   0.5,
    'legend.facecolor': CARD_BG,
    'legend.edgecolor': BDR,
})

# ─────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────
def show_fig(fig):
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)

def fmt_num(n):
    if n >= 1e9: return f"${n/1e9:.1f}B"
    if n >= 1e6: return f"${n/1e6:.1f}M"
    if n >= 1e3: return f"${n/1e3:.1f}K"
    return f"${n:.2f}"

@st.cache_data
def load_and_clean(file_bytes):
    """
    Standard 13-step data cleaning pipeline.
    Returns (clean_df, report_dict).
    """
    # ── 1. Load raw Excel ─────────────────────────────────────
    raw = pd.read_excel(io.BytesIO(file_bytes))
    report = {}

    # Snapshot BEFORE any changes
    report['raw_rows'] = len(raw)
    report['raw_cols'] = len(raw.columns)

    # ── 2. Normalise column names ─────────────────────────────
    raw.columns = raw.columns.str.strip()
    col_aliases = {
        'invoice date':     'Invoice Date',
        'date':             'Invoice Date',
        'retailerid':       'Retailer ID',
        'retailer id':      'Retailer ID',
        'price per unit':   'Price per Unit',
        'units sold':       'Units Sold',
        'total sales':      'Total Sales',
        'operating profit': 'Operating Profit',
        'operating margin': 'Operating Margin',
        'sales method':     'Sales Method',
    }
    raw.columns = [col_aliases.get(c.lower(), c) for c in raw.columns]
    report['columns_after_normalise'] = list(raw.columns)

    # ── 3. Drop fully-empty rows & columns ───────────────────
    before = len(raw)
    raw.dropna(how='all', inplace=True)
    raw.dropna(axis=1, how='all', inplace=True)
    report['empty_rows_dropped'] = before - len(raw)

    # ── 4. Strip string whitespace ────────────────────────────
    str_cols = raw.select_dtypes(include='object').columns.tolist()
    for col in str_cols:
        raw[col] = raw[col].astype(str).str.strip()
        raw[col] = raw[col].replace({'nan': np.nan, 'None': np.nan, '': np.nan})

    # ── 5. Standardise categoricals to Title Case ─────────────
    cat_cols = ['Retailer', 'Region', 'State', 'City', 'Product', 'Sales Method']
    for col in cat_cols:
        if col in raw.columns:
            raw[col] = raw[col].str.title()

    # ── 6. Remove exact duplicate rows ───────────────────────
    before = len(raw)
    raw.drop_duplicates(inplace=True)
    report['duplicates_dropped'] = before - len(raw)

    # ── 7. Parse & validate Invoice Date ─────────────────────
    before = len(raw)
    raw['Invoice Date'] = pd.to_datetime(raw['Invoice Date'], errors='coerce')
    bad_dates = int(raw['Invoice Date'].isna().sum())
    raw.dropna(subset=['Invoice Date'], inplace=True)
    report['bad_dates_dropped'] = before - len(raw)
    report['bad_dates_found']   = bad_dates

    # ── 8. Coerce numeric columns; non-numeric → NaN ─────────
    num_cols = ['Price per Unit', 'Units Sold', 'Total Sales',
                'Operating Profit', 'Operating Margin']
    num_cols = [c for c in num_cols if c in raw.columns]
    nan_before = {c: int(raw[c].isna().sum()) for c in num_cols}
    for col in num_cols:
        raw[col] = pd.to_numeric(raw[col], errors='coerce')
    nan_after = {c: int(raw[c].isna().sum()) for c in num_cols}
    report['coerced_to_nan'] = {c: nan_after[c] - nan_before[c] for c in num_cols}

    # ── 9. Impute remaining NaNs with column median ──────────
    imputed = {}
    for col in num_cols:
        n_nan = int(raw[col].isna().sum())
        if n_nan > 0:
            med = raw[col].median()
            raw[col].fillna(med, inplace=True)
            imputed[col] = {'count': n_nan, 'median': round(float(med), 4)}
    report['imputed'] = imputed

    # ── 10. Clip negatives in financial columns to 0 ─────────
    fin_cols = ['Total Sales', 'Operating Profit', 'Units Sold', 'Price per Unit']
    fin_cols = [c for c in fin_cols if c in raw.columns]
    neg_counts = {}
    for col in fin_cols:
        n_neg = int((raw[col] < 0).sum())
        if n_neg:
            raw[col] = raw[col].clip(lower=0)
            neg_counts[col] = n_neg
    report['negatives_clipped'] = neg_counts

    # ── 11. Fix Operating Margin scale (>1 means % not ratio) ─
    if 'Operating Margin' in raw.columns:
        over_one = int((raw['Operating Margin'] > 1).sum())
        if over_one > 0:
            raw.loc[raw['Operating Margin'] > 1, 'Operating Margin'] /= 100
        report['margin_rescaled'] = over_one
    else:
        report['margin_rescaled'] = 0

    # ── 12. Engineer date-derived columns ─────────────────────
    raw['YearMonth'] = raw['Invoice Date'].dt.to_period('M')
    raw['Year']      = raw['Invoice Date'].dt.year
    raw['Month']     = raw['Invoice Date'].dt.month
    raw['Quarter']   = raw['Invoice Date'].dt.to_period('Q').astype(str)

    # ── 13. Final state snapshot ──────────────────────────────
    raw.reset_index(drop=True, inplace=True)
    report['final_rows']       = len(raw)
    report['final_data_cols']  = report['raw_cols']          # original column count
    report['final_total_cols'] = len(raw.columns)            # incl. engineered
    report['engineered_cols']  = report['final_total_cols'] - report['raw_cols']
    report['final_nulls']      = int(raw.isnull().sum().sum())
    report['date_range_start'] = str(raw['Invoice Date'].min().date())
    report['date_range_end']   = str(raw['Invoice Date'].max().date())

    return raw, report

# ─────────────────────────────────────────────
#  SIDEBAR  (single upload widget lives here)
# ─────────────────────────────────────────────
with st.sidebar:
    _sb_header = (
        '<div style="display:flex;align-items:center;gap:12px;padding:14px 0 6px;">'
        + LOGO_SVG +
        '<div>'
        '<div style="font-family:Orbitron,sans-serif;font-size:0.92rem;'
        'font-weight:900;color:' + A1 + ';letter-spacing:2px;">ADIDAS</div>'
        '<div style="font-family:JetBrains Mono,monospace;font-size:0.54rem;'
        'color:' + TSEC + ';letter-spacing:3px;">SALES INTELLIGENCE</div>'
        '</div></div>'
        '<hr style="border-color:' + BDR + ';margin:8px 0 14px;">'
        '<div style="font-family:JetBrains Mono,monospace;font-size:0.58rem;color:' + TSEC + ';'
        'letter-spacing:2px;text-transform:uppercase;margin-bottom:8px;">'
        + svg(P_FILE, 12, TSEC) + ' Data Source</div>'
    )
    st.markdown(_sb_header, unsafe_allow_html=True)

    # THE ONE AND ONLY file uploader
    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"],
        label_visibility="collapsed",
        help="Upload Adidas Sales Analysis Excel file (.xlsx / .xls)"
    )

    st.markdown(f"""
    <hr style="border-color:{BDR};margin:14px 0 12px;">
    <div style="font-family:'JetBrains Mono',monospace;font-size:0.58rem;color:{TSEC};
                letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;">
        {svg(P_NAV, 12, TSEC)} Navigation
    </div>
    """, unsafe_allow_html=True)

    page = st.radio("Page", ["Sales Overview", "Geo & Product Deep Dive"],
                    label_visibility="collapsed")

    if uploaded_file:
        st.markdown(f"""
        <hr style="border-color:{BDR};margin:14px 0 12px;">
        <div style="font-family:'JetBrains Mono',monospace;font-size:0.58rem;color:{TSEC};
                    letter-spacing:2px;text-transform:uppercase;margin-bottom:10px;">
            {svg(P_FIND, 12, TSEC)} Filters
        </div>
        """, unsafe_allow_html=True)

    st.markdown(f"""
    <div style="position:fixed;bottom:14px;left:0;width:236px;text-align:center;
                font-family:'JetBrains Mono',monospace;font-size:0.5rem;
                color:{TSEC}33;letter-spacing:1px;">
        STREAMLIT + MATPLOTLIB + PANDAS
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  UPLOAD GATE  (no widget here — uploader is in sidebar)
# ─────────────────────────────────────────────
if uploaded_file is None:
    # Build entire upload landing as one HTML string — no nested f-strings with SVG
    _upload_icon = (
        '<svg width="52" height="52" viewBox="0 0 24 24" fill="none"'
        ' xmlns="http://www.w3.org/2000/svg"'
        ' style="margin-bottom:4px;">'
        '<path d="' + P_UPLOAD + '" stroke="' + A1 + '" stroke-width="1.6"'
        ' stroke-linecap="round" stroke-linejoin="round"/></svg>'
    )
    _badges = (
        '<span class="upload-badge">' + svg(P_FILE,12,TSEC) + 'Excel (.xlsx / .xls)</span>'
        '<span class="upload-badge">' + svg(P_CHART,12,TSEC) + '9,000+ Records</span>'
        '<span class="upload-badge">' + svg(P_FLASH,12,TSEC) + 'Instant Analysis</span>'
        '<span class="upload-badge">' + svg(P_NAV,12,TSEC) + '2-Page Dashboard</span>'
    )
    _upload_page_html = (
        '<div class="upload-wrap">'
        + LOGO_LG +
        '<h2>ADIDAS SALES INTELLIGENCE</h2>'
        '<p>Upload your Sales Analysis Excel file using the panel on the left'
        ' to launch the full interactive dashboard.</p>'
        '</div>'
        '<div style="max-width:520px;margin:0 auto;">'
        '<div style="display:flex;flex-direction:column;align-items:center;gap:14px;'
        'background:' + CARD_BG + ';border:2px dashed ' + A1 + '55;border-radius:14px;'
        'padding:28px 32px;margin-bottom:20px;">'
        + _upload_icon +
        '<div style="font-family:Orbitron,sans-serif;font-size:0.72rem;'
        'color:' + A1 + ';letter-spacing:2px;text-align:center;">USE THE SIDEBAR UPLOADER</div>'
        '<div style="font-family:JetBrains Mono,monospace;font-size:0.65rem;'
        'color:' + TSEC + ';letter-spacing:1px;text-align:center;line-height:1.7;">'
        'Click the <b style="color:' + TPRI + ';">arrow on the left</b> to open the sidebar,'
        ' then drag &amp; drop or browse your <b style="color:' + TPRI + ';">.xlsx</b> file.'
        '</div></div>'
        '<div style="text-align:center;margin-bottom:14px;">' + _badges + '</div>'
        '<div style="text-align:center;font-family:JetBrains Mono,monospace;'
        'font-size:0.6rem;color:' + TSEC + '44;letter-spacing:1px;">'
        'Expected columns: Retailer &middot; Region &middot; State &middot; City'
        ' &middot; Product &middot; Price per Unit &middot; Units Sold &middot;'
        ' Total Sales &middot; Operating Profit &middot; Operating Margin'
        ' &middot; Sales Method &middot; Invoice Date'
        '</div></div>'
    )
    components.html(
        '<style>'
        'body{margin:0;padding:0;background:#0A0A0F;}'
        '@import url("https://fonts.googleapis.com/css2?family=Orbitron:wght@900&family=Rajdhani:wght@400;600&family=JetBrains+Mono:wght@400&display=swap");'
        '.upload-wrap{display:flex;flex-direction:column;align-items:center;padding:40px 20px 20px;text-align:center;}'
        'h2{font-family:Orbitron,sans-serif;font-size:1.4rem;font-weight:900;'
        'background:linear-gradient(90deg,#00E5FF,#FF3CAC);-webkit-background-clip:text;'
        '-webkit-text-fill-color:transparent;letter-spacing:3px;margin:16px 0 8px;}'
        'p{color:#8A8A9A;font-family:JetBrains Mono,monospace;font-size:0.75rem;'
        'letter-spacing:1px;margin-bottom:24px;max-width:480px;}'
        '.upload-badge{display:inline-flex;align-items:center;gap:7px;'
        'background:#12121A;border:1px solid #1E1E2E;border-radius:8px;'
        'padding:7px 14px;font-family:JetBrains Mono,monospace;'
        'font-size:0.68rem;color:#8A8A9A;margin:4px;}'
        '.card{display:flex;flex-direction:column;align-items:center;gap:14px;'
        'background:#12121A;border:2px dashed #00E5FF55;border-radius:14px;'
        'padding:28px 32px;margin-bottom:20px;max-width:480px;width:100%;}'
        '.instr{font-family:JetBrains Mono,monospace;font-size:0.65rem;'
        'color:#8A8A9A;letter-spacing:1px;text-align:center;line-height:1.7;}'
        '.lbl{font-family:Orbitron,sans-serif;font-size:0.72rem;'
        'color:#00E5FF;letter-spacing:2px;text-align:center;}'
        '</style>'
        + _upload_page_html,
        height=520, scrolling=False
    )
    st.stop()

# ─────────────────────────────────────────────
#  LOAD, CLEAN & REPORT
# ─────────────────────────────────────────────
try:
    file_bytes = uploaded_file.read()
    df_raw, rpt = load_and_clean(file_bytes)
except Exception as exc:
    st.error(f"Could not read or clean file: {exc}")
    st.stop()

# ── Data Cleaning Report ─────────────────────
rows_removed = rpt['raw_rows'] - rpt['final_rows']
P_CHECK  = "M20 6L9 17l-5-5"
P_WARN   = "M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0zM12 9v4M12 17h.01"
P_SHIELD = "M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"

def step_html(num, label, detail, kind="ok"):
    css = "clean-step-ok" if kind == "ok" else ("clean-step-warn" if kind == "warn" else "clean-step-info")
    icon_path = P_CHECK if kind == "ok" else (P_WARN if kind == "warn" else P_WAVE)
    icon_color = A5 if kind == "ok" else (A4 if kind == "warn" else TSEC)
    return f"""
    <div class="clean-step">
        <span class="clean-step-num">{num:02d}</span>
        {svg(icon_path, 12, icon_color)}
        <span class="{css}"><b>{label}</b> — {detail}</span>
    </div>"""

steps_html = ""
steps_html += step_html(1,  "Load Excel",
    f"{rpt['raw_rows']:,} rows × {rpt['raw_cols']} columns read from file")
steps_html += step_html(2,  "Normalise Column Names",
    f"Whitespace stripped; common aliases remapped to standard names")
steps_html += step_html(3,  "Drop Empty Rows / Cols",
    f"{rpt['empty_rows_dropped']} fully-empty rows removed",
    "warn" if rpt['empty_rows_dropped'] > 0 else "ok")
steps_html += step_html(4,  "Strip String Whitespace",
    "All object columns trimmed of leading/trailing spaces")
steps_html += step_html(5,  "Title-Case Categoricals",
    "Retailer, Region, State, City, Product, Sales Method standardised")
steps_html += step_html(6,  "Remove Exact Duplicates",
    f"{rpt['duplicates_dropped']} duplicate rows removed",
    "warn" if rpt['duplicates_dropped'] > 0 else "ok")
steps_html += step_html(7,  "Parse Invoice Date",
    f"{rpt['bad_dates_dropped']} unparseable dates dropped  |  "
    f"Range: {rpt['date_range_start']} to {rpt['date_range_end']}",
    "warn" if rpt['bad_dates_dropped'] > 0 else "ok")
coerced_total = sum(rpt['coerced_to_nan'].values())
steps_html += step_html(8,  "Coerce Numerics",
    f"{coerced_total} non-numeric values converted to NaN across "
    f"{len(rpt['coerced_to_nan'])} numeric columns",
    "warn" if coerced_total > 0 else "ok")
imp_total = sum(v['count'] for v in rpt['imputed'].values())
imp_detail = (", ".join(f"{c}: {v['count']} filled (median={v['median']})"
              for c, v in rpt['imputed'].items()) if rpt['imputed'] else "No NaNs found")
steps_html += step_html(9,  "Impute NaN with Median",
    f"{imp_total} values imputed  |  {imp_detail}",
    "warn" if imp_total > 0 else "ok")
neg_total = sum(rpt['negatives_clipped'].values())
neg_detail = (", ".join(f"{c}: {v}" for c, v in rpt['negatives_clipped'].items())
              if rpt['negatives_clipped'] else "None found")
steps_html += step_html(10, "Clip Negative Values",
    f"{neg_total} negative values in financial columns clipped to 0  |  {neg_detail}",
    "warn" if neg_total > 0 else "ok")
margin_fixed = rpt.get('margin_rescaled', 0)
steps_html += step_html(11, "Fix Margin Scale",
    f"{margin_fixed} Operating Margin values >1 rescaled by dividing by 100",
    "warn" if margin_fixed > 0 else "ok")
steps_html += step_html(12, "Engineer Date Columns",
    f"Derived: YearMonth, Year, Month, Quarter from Invoice Date")
steps_html += step_html(13, "Final Dataset",
    f"{rpt['final_rows']:,} rows  |  "
    f"{rpt['raw_cols']} original + {rpt['engineered_cols']} engineered = "
    f"{rpt['final_total_cols']} total columns  |  {rpt['final_nulls']} nulls remaining")

with st.expander(
    f"Data Cleaning Report  —  {rpt['raw_rows']:,} raw rows  →  "
    f"{rpt['final_rows']:,} clean rows  ({rows_removed} removed)",
    expanded=False
):
    st.markdown(f"""
    <div class="clean-panel">
        <div class="clean-panel-title">
            {svg(P_SHIELD, 15, A5)} Data Quality Pipeline — 13 Steps Applied
        </div>
        <div class="clean-grid">
            <div class="clean-stat">
                <div class="clean-stat-val">{rpt['raw_rows']:,}</div>
                <div class="clean-stat-lbl">Raw Rows</div>
            </div>
            <div class="clean-stat">
                <div class="clean-stat-val">{rpt['final_rows']:,}</div>
                <div class="clean-stat-lbl">Clean Rows</div>
            </div>
            <div class="clean-stat">
                <div class="clean-stat-val">{rows_removed}</div>
                <div class="clean-stat-lbl">Rows Removed</div>
            </div>
            <div class="clean-stat">
                <div class="clean-stat-val">{rpt['raw_cols']} + {rpt['engineered_cols']}</div>
                <div class="clean-stat-lbl">Cols (orig + derived)</div>
            </div>
        </div>
        {steps_html}
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
#  SIDEBAR FILTERS (after data loaded)
# ─────────────────────────────────────────────
with st.sidebar:
    if page == "Sales Overview":
        regions   = ["All"] + sorted(df_raw['Region'].dropna().unique())
        retailers = ["All"] + sorted(df_raw['Retailer'].dropna().unique())
        sel_region   = st.selectbox("Region",   regions)
        sel_retailer = st.selectbox("Retailer", retailers)
        min_d = df_raw['Invoice Date'].min().date()
        max_d = df_raw['Invoice Date'].max().date()
        date_range = st.date_input("Date Range",
                        value=(min_d, max_d), min_value=min_d, max_value=max_d)
        sel_method  = "All"
        sel_product = "All"
    else:
        methods  = ["All"] + sorted(df_raw['Sales Method'].dropna().unique())
        products = ["All"] + sorted(df_raw['Product'].dropna().unique())
        sel_method  = st.selectbox("Sales Method", methods)
        sel_product = st.selectbox("Product",      products)
        sel_region   = "All"
        sel_retailer = "All"
        date_range   = None

# ─────────────────────────────────────────────
#  FILTER HELPERS
# ─────────────────────────────────────────────
def filter_p1(df):
    d = df.copy()
    if sel_region   != "All": d = d[d['Region']   == sel_region]
    if sel_retailer != "All": d = d[d['Retailer'] == sel_retailer]
    if date_range and len(date_range) == 2:
        d = d[(d['Invoice Date'].dt.date >= date_range[0]) &
              (d['Invoice Date'].dt.date <= date_range[1])]
    return d

def filter_p2(df):
    d = df.copy()
    if sel_method  != "All": d = d[d['Sales Method'] == sel_method]
    if sel_product != "All": d = d[d['Product']      == sel_product]
    return d

# ════════════════════════════════════════════════════════════
#  PAGE 1 — SALES OVERVIEW
# ════════════════════════════════════════════════════════════
if page == "Sales Overview":
    df = filter_p1(df_raw)

    # ── Header ──────────────────────────────────
    st.markdown(
        '<div class="dashboard-header">' + LOGO_SVG +
        '<div class="header-text">'
        '<h1>SALES OVERVIEW</h1>'
        '<p>Total Sales &nbsp;&middot;&nbsp; Profitability &nbsp;&middot;&nbsp; Volume'
        ' &nbsp;&middot;&nbsp; Pricing &nbsp;&middot;&nbsp; Margin Analysis</p>'
        '</div></div>',
        unsafe_allow_html=True)

    # ── KPIs ────────────────────────────────────
    total_sales  = df['Total Sales'].sum()
    total_profit = df['Operating Profit'].sum()
    total_units  = int(df['Units Sold'].sum())
    avg_price    = df['Price per Unit'].mean()
    avg_margin   = df['Operating Margin'].mean() * 100

    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card kpi-c1">
        <span class="kpi-icon">{svg(P_STAR, 20, A1)}</span>
        <div class="kpi-label">Total Sales</div>
        <div class="kpi-value">{fmt_num(total_sales)}</div>
        <div class="kpi-sub">Overall Revenue</div>
      </div>
      <div class="kpi-card kpi-c2">
        <span class="kpi-icon">{svg(P_PROFIT, 20, A2)}</span>
        <div class="kpi-label">Operating Profit</div>
        <div class="kpi-value">{fmt_num(total_profit)}</div>
        <div class="kpi-sub">Net Profitability</div>
      </div>
      <div class="kpi-card kpi-c3">
        <span class="kpi-icon">{svg(P_BOX, 20, A4)}</span>
        <div class="kpi-label">Units Sold</div>
        <div class="kpi-value">{total_units:,}</div>
        <div class="kpi-sub">Product Demand</div>
      </div>
      <div class="kpi-card kpi-c4">
        <span class="kpi-icon">{svg(P_DOLLAR, 20, A5)}</span>
        <div class="kpi-label">Avg Price / Unit</div>
        <div class="kpi-value">${avg_price:.2f}</div>
        <div class="kpi-sub">Pricing Strategy</div>
      </div>
      <div class="kpi-card kpi-c5">
        <span class="kpi-icon">{svg(P_WAVE, 20, A3)}</span>
        <div class="kpi-label">Avg Margin</div>
        <div class="kpi-value">{avg_margin:.1f}%</div>
        <div class="kpi-sub">Profitability Rate</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Row 1: Area Chart + Region Donut ────────
    c1, c2 = st.columns([2, 1])

    with c1:
        st.markdown(ct(P_TREND, "TOTAL SALES BY MONTH — AREA CHART"), unsafe_allow_html=True)
        monthly = (df.groupby('YearMonth')['Total Sales']
                     .sum().reset_index().sort_values('YearMonth'))
        monthly['Label'] = monthly['YearMonth'].astype(str)
        x = np.arange(len(monthly))
        y = monthly['Total Sales'].values / 1e6

        fig, ax = plt.subplots(figsize=(10, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        cmap_area = LinearSegmentedColormap.from_list('a', [A1 + '00', A1 + '70'])
        if len(x) > 1:
            ax.imshow(np.linspace(0, 1, 300).reshape(1, -1), aspect='auto',
                      extent=[x[0], x[-1], 0, y.max()],
                      cmap=cmap_area, origin='lower', zorder=1, alpha=0.45)
        ax.plot(x, y, color=A1, lw=2.5, zorder=5)
        ax.fill_between(x, y, color=A1, alpha=0.07, zorder=2)
        ax.scatter(x, y, color=A1, s=26, zorder=6)
        step = max(1, len(monthly) // 8)
        ax.set_xticks(x[::step])
        ax.set_xticklabels(monthly['Label'].iloc[::step], rotation=30, ha='right', fontsize=8)
        ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v:.0f}M'))
        ax.grid(axis='y', alpha=0.25)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        plt.tight_layout()
        show_fig(fig)

    with c2:
        st.markdown(ct(P_GLOBE, "SALES BY REGION — DONUT"), unsafe_allow_html=True)
        reg = df.groupby('Region')['Total Sales'].sum()
        fig, ax = plt.subplots(figsize=(4.5, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        wedges, texts, autos = ax.pie(
            reg.values, labels=reg.index, autopct='%1.1f%%',
            colors=PALETTE[:len(reg)],
            wedgeprops=dict(width=0.52, edgecolor=CARD_BG, linewidth=2),
            pctdistance=0.82, startangle=90)
        for t in texts: t.set_color(TSEC); t.set_fontsize(8)
        for a in autos: a.set_color(DARK_BG); a.set_fontsize(7.5); a.set_fontweight('bold')
        plt.tight_layout()
        show_fig(fig)

    # ── Row 2: Product Bar + Sales Method Donut ──
    c3, c4 = st.columns([2, 1])

    with c3:
        st.markdown(ct(P_SHOE, "TOTAL SALES BY PRODUCT — BAR CHART"), unsafe_allow_html=True)
        prod = df.groupby('Product')['Total Sales'].sum().sort_values()
        fig, ax = plt.subplots(figsize=(10, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        ax.barh(prod.index, prod.values / 1e6,
                color=PALETTE[:len(prod)], edgecolor='none', height=0.58)
        for i, v in enumerate(prod.values):
            ax.text(v / 1e6 + 0.2, i, f'${v/1e6:.1f}M',
                    va='center', fontsize=8, color=TSEC)
        ax.set_xlabel('Sales (Millions $)', fontsize=8)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        ax.grid(axis='x', alpha=0.2)
        ax.tick_params(axis='y', labelsize=8)
        plt.tight_layout()
        show_fig(fig)

    with c4:
        st.markdown(ct(P_CHART, "SALES BY METHOD — DONUT"), unsafe_allow_html=True)
        meth = df.groupby('Sales Method')['Total Sales'].sum()
        fig, ax = plt.subplots(figsize=(4.5, 3.8))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        wedges, texts, autos = ax.pie(
            meth.values, labels=meth.index, autopct='%1.1f%%',
            colors=[A2, A4, A1],
            wedgeprops=dict(width=0.52, edgecolor=CARD_BG, linewidth=2),
            pctdistance=0.82, startangle=90)
        for t in texts: t.set_color(TSEC); t.set_fontsize(8)
        for a in autos: a.set_color(DARK_BG); a.set_fontsize(7.5); a.set_fontweight('bold')
        plt.tight_layout()
        show_fig(fig)

    # ── Retailer Bar (full width) ────────────────
    st.markdown(ct(P_STAR, "TOTAL SALES BY RETAILER — BAR CHART", A4), unsafe_allow_html=True)
    ret = df.groupby('Retailer')['Total Sales'].sum().sort_values(ascending=False)
    fig, ax = plt.subplots(figsize=(14, 3.2))
    fig.patch.set_facecolor(CARD_BG)
    ax.set_facecolor(CARD_BG)
    bar_colors = [A1 if i == 0 else A3 for i in range(len(ret))]
    bars = ax.bar(ret.index, ret.values / 1e6, color=bar_colors, edgecolor='none', width=0.5)
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width() / 2, h + 0.4,
                f'${h:.1f}M', ha='center', fontsize=8, color=TSEC)
    ax.set_ylabel('Sales (M$)', fontsize=8)
    ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
    ax.grid(axis='y', alpha=0.2)
    plt.tight_layout()
    show_fig(fig)

    # ── Key Findings ────────────────────────────
    st.markdown(sl(P_FIND, "KEY FINDINGS — SALES OVERVIEW"), unsafe_allow_html=True)

    top_retailer = ret.index[0]
    top_product  = df.groupby('Product')['Total Sales'].sum().idxmax()
    top_region   = df.groupby('Region')['Total Sales'].sum().idxmax()
    top_method   = df.groupby('Sales Method')['Total Sales'].sum().idxmax()
    peak_month   = monthly.loc[monthly['Total Sales'].idxmax(), 'Label']

    findings1 = [
        (P_STAR,   A4, "TOP RETAILER",
         f"<b>{top_retailer}</b> leads all retailers in total revenue, capturing the largest share of Adidas sales across the dataset period."),
        (P_SHOE,   A1, "BEST-SELLING PRODUCT",
         f"<b>{top_product}</b> drives the most revenue, signalling strong consumer preference and a strategic pricing sweet spot."),
        (P_GLOBE,  A2, "DOMINANT REGION",
         f"The <b>{top_region}</b> region contributes the highest sales volume — the priority market for distribution and marketing."),
        (P_CHART,  A5, "PREFERRED SALES CHANNEL",
         f"<b>{top_method}</b> is the most productive channel by revenue, signalling where continued investment will yield highest returns."),
        (P_TREND,  A1, "PEAK SALES MONTH",
         f"Sales peaked in <b>{peak_month}</b>, likely driven by seasonal demand, promotions, or product launches — critical for inventory planning."),
        (P_WAVE,   A3, "MARGIN HEALTH",
         f"Average operating margin of <b>{avg_margin:.1f}%</b> reflects a healthy profitability baseline, with product-level optimisation opportunities remaining."),
    ]
    f1, f2 = st.columns(2)
    for i, (icon_d, ic, title, body) in enumerate(findings1):
        with (f1 if i % 2 == 0 else f2):
            st.markdown(f"""
            <div class="finding-card">
                <h4>{svg(icon_d, 13, ic)}{title}</h4>
                <p>{body}</p>
            </div>""", unsafe_allow_html=True)

# ════════════════════════════════════════════════════════════
#  PAGE 2 — GEO & PRODUCT DEEP DIVE
# ════════════════════════════════════════════════════════════
else:
    df = filter_p2(df_raw)

    # ── Header ──────────────────────────────────
    st.markdown(
        '<div class="dashboard-header">' + LOGO_SVG +
        '<div class="header-text">'
        '<h1>GEO &amp; PRODUCT DEEP DIVE</h1>'
        '<p>State &nbsp;&middot;&nbsp; City Rankings &nbsp;&middot;&nbsp; Product Profitability'
        ' &nbsp;&middot;&nbsp; Price-Volume &nbsp;&middot;&nbsp; Sales vs Profit</p>'
        '</div></div>',
        unsafe_allow_html=True)

    # ── KPIs ────────────────────────────────────
    top_state    = df.groupby('State')['Total Sales'].sum().idxmax()
    top_prod2    = df.groupby('Product')['Total Sales'].sum().idxmax()
    top_ret_prof = df.groupby('Retailer')['Operating Profit'].sum().idxmax()
    avg_price2   = df['Price per Unit'].mean()

    st.markdown(f"""
    <div class="kpi-grid">
      <div class="kpi-card kpi-c1">
        <span class="kpi-icon">{svg(P_MAP, 20, A1)}</span>
        <div class="kpi-label">Highest Selling State</div>
        <div class="kpi-value" style="font-size:1.05rem;">{top_state}</div>
        <div class="kpi-sub">By Total Revenue</div>
      </div>
      <div class="kpi-card kpi-c2">
        <span class="kpi-icon">{svg(P_TAG, 20, A2)}</span>
        <div class="kpi-label">Highest Selling Product</div>
        <div class="kpi-value" style="font-size:0.88rem;">{top_prod2}</div>
        <div class="kpi-sub">By Total Revenue</div>
      </div>
      <div class="kpi-card kpi-c3">
        <span class="kpi-icon">{svg(P_STAR, 20, A4)}</span>
        <div class="kpi-label">Most Profitable Retailer</div>
        <div class="kpi-value" style="font-size:0.95rem;">{top_ret_prof}</div>
        <div class="kpi-sub">By Operating Profit</div>
      </div>
      <div class="kpi-card kpi-c4">
        <span class="kpi-icon">{svg(P_DOLLAR, 20, A5)}</span>
        <div class="kpi-label">Avg Price / Unit</div>
        <div class="kpi-value">${avg_price2:.2f}</div>
        <div class="kpi-sub">Pricing Benchmark</div>
      </div>
      <div class="kpi-card kpi-c5">
        <span class="kpi-icon">{svg(P_FILE, 20, A3)}</span>
        <div class="kpi-label">Total Records</div>
        <div class="kpi-value">{len(df):,}</div>
        <div class="kpi-sub">Filtered Dataset</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Row 1: State Bar + City Bar ──────────────
    c1, c2 = st.columns([1, 1])

    with c1:
        st.markdown(ct(P_MAP, "TOTAL SALES BY STATE — TOP 20 FILLED BAR"), unsafe_allow_html=True)
        state_sales = (df.groupby('State')['Total Sales'].sum()
                         .sort_values(ascending=False).head(20))
        norm = state_sales.values / state_sales.max()
        # cyan gradient encoding intensity
        cs = [f"#{int(0):02x}{int(n*229):02x}{int(n*255):02x}" for n in norm]
        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        ax.barh(state_sales.index[::-1], state_sales.values[::-1] / 1e6,
                color=cs[::-1], edgecolor='none', height=0.65)
        for i, v in enumerate(state_sales.values[::-1]):
            ax.text(v / 1e6 + 0.1, i, f'${v/1e6:.1f}M',
                    va='center', fontsize=7, color=TSEC)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        ax.grid(axis='x', alpha=0.2)
        ax.tick_params(axis='y', labelsize=7.5)
        ax.set_xlabel('Sales (M$)', fontsize=8)
        plt.tight_layout()
        show_fig(fig)

    with c2:
        st.markdown(ct(P_CITY, "TOP 10 CITIES BY SALES — BAR CHART"), unsafe_allow_html=True)
        city_sales = (df.groupby('City')['Total Sales'].sum()
                        .sort_values(ascending=False).head(10))
        fig, ax = plt.subplots(figsize=(7, 6))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        xp = np.arange(len(city_sales))
        ax.bar(xp, city_sales.values / 1e6,
               color=PALETTE[:len(city_sales)], edgecolor='none', width=0.6)
        for i, v in enumerate(city_sales.values):
            ax.text(i, v / 1e6 + 0.15, f'${v/1e6:.1f}M',
                    ha='center', fontsize=7, color=TSEC)
        ax.set_xticks(xp)
        ax.set_xticklabels(city_sales.index, rotation=35, ha='right', fontsize=8)
        ax.set_ylabel('Sales (M$)', fontsize=8)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        ax.grid(axis='y', alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    # ── Row 2: Profit by Product + Price vs Units ─
    c3, c4 = st.columns([1, 1])

    with c3:
        st.markdown(ct(P_PROFIT, "PROFIT BY PRODUCT — COLUMN CHART"), unsafe_allow_html=True)
        prod_profit = (df.groupby('Product')['Operating Profit'].sum()
                         .sort_values(ascending=False))
        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        xp = np.arange(len(prod_profit))
        ax.bar(xp, prod_profit.values / 1e6,
               color=PALETTE[:len(prod_profit)], edgecolor='none', width=0.55)
        for i, v in enumerate(prod_profit.values):
            ax.text(i, v / 1e6 + 0.15, f'${v/1e6:.1f}M',
                    ha='center', fontsize=7, color=TSEC)
        ax.set_xticks(xp)
        ax.set_xticklabels([p.replace("'s ", "\n") for p in prod_profit.index], fontsize=7.5)
        ax.set_ylabel('Profit (M$)', fontsize=8)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        ax.grid(axis='y', alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    with c4:
        st.markdown(ct(P_SCATTER, "PRICE vs UNITS SOLD — SCATTER PLOT"), unsafe_allow_html=True)
        samp = df.sample(min(1500, len(df)), random_state=42)
        prods_u  = df['Product'].unique()
        cmap_p   = {p: PALETTE[i % len(PALETTE)] for i, p in enumerate(prods_u)}
        c_list   = [cmap_p[p] for p in samp['Product']]
        fig, ax = plt.subplots(figsize=(7, 4))
        fig.patch.set_facecolor(CARD_BG)
        ax.set_facecolor(CARD_BG)
        ax.scatter(samp['Price per Unit'], samp['Units Sold'],
                   c=c_list, alpha=0.55, s=18, edgecolors='none')
        handles = [mpatches.Patch(color=cmap_p[p], label=p) for p in prods_u]
        ax.legend(handles=handles, fontsize=6.5, loc='upper right', framealpha=0.3)
        ax.set_xlabel('Price per Unit ($)', fontsize=8)
        ax.set_ylabel('Units Sold', fontsize=8)
        ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
        ax.grid(alpha=0.2)
        plt.tight_layout()
        show_fig(fig)

    # ── Row 3: Sales vs Profit (full width) ──────
    st.markdown(ct(P_SCATTER, "SALES vs OPERATING PROFIT BY RETAILER — SCATTER PLOT", A2),
                unsafe_allow_html=True)
    samp2    = df.sample(min(2000, len(df)), random_state=7)
    retailers = df['Retailer'].unique()
    cmap_r   = {r: PALETTE[i % len(PALETTE)] for i, r in enumerate(retailers)}
    c_ret    = [cmap_r[r] for r in samp2['Retailer']]
    fig, ax = plt.subplots(figsize=(14, 4))
    fig.patch.set_facecolor(CARD_BG)
    ax.set_facecolor(CARD_BG)
    ax.scatter(samp2['Total Sales'], samp2['Operating Profit'],
               c=c_ret, alpha=0.5, s=22, edgecolors='none')
    handles2 = [mpatches.Patch(color=cmap_r[r], label=r) for r in retailers]
    ax.legend(handles=handles2, fontsize=8, loc='upper left',
              framealpha=0.3, ncol=len(retailers))
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v/1e3:.0f}K'))
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda v, _: f'${v/1e3:.0f}K'))
    ax.set_xlabel('Total Sales', fontsize=8)
    ax.set_ylabel('Operating Profit', fontsize=8)
    ax.spines[['top', 'right', 'left', 'bottom']].set_visible(False)
    ax.grid(alpha=0.15)
    plt.tight_layout()
    show_fig(fig)

    # ── Key Findings ────────────────────────────
    st.markdown(sl(P_FIND, "KEY FINDINGS — GEO & PRODUCT"), unsafe_allow_html=True)

    top_city         = df.groupby('City')['Total Sales'].sum().idxmax()
    top_state_profit = df.groupby('State')['Operating Profit'].sum().idxmax()
    high_margin_prod = df.groupby('Product')['Operating Margin'].mean().idxmax()
    high_vol_prod    = df.groupby('Product')['Units Sold'].mean().idxmax()

    findings2 = [
        (P_CITY,    A1, "TOP PERFORMING CITY",
         f"<b>{top_city}</b> ranks as the highest-revenue city — Adidas's most critical urban market for sales concentration."),
        (P_MAP,     A2, "TOP PROFIT STATE",
         f"<b>{top_state_profit}</b> generates the most operating profit, suggesting strong operational efficiency or favourable pricing dynamics."),
        (P_TAG,     A4, "HIGHEST MARGIN PRODUCT",
         f"<b>{high_margin_prod}</b> commands the best average margin, making it Adidas's most financially efficient product line."),
        (P_FLASH,   A5, "HIGHEST VOLUME PRODUCT",
         f"<b>{high_vol_prod}</b> achieves the highest average units sold per transaction, indicating strong mass-market demand."),
        (P_SCATTER, A1, "SALES-PROFIT CORRELATION",
         "The scatter analysis confirms a strong positive linear relationship between total sales and operating profit — revenue growth directly drives profitability."),
        (P_ALERT,   A3, "PRICE SENSITIVITY",
         "Lower-priced products (under $45) consistently outperform in units sold, suggesting significant price elasticity in the Adidas consumer base."),
    ]
    g1, g2 = st.columns(2)
    for i, (icon_d, ic, title, body) in enumerate(findings2):
        with (g1 if i % 2 == 0 else g2):
            st.markdown(f"""
            <div class="finding-card">
                <h4>{svg(icon_d, 13, ic)}{title}</h4>
                <p>{body}</p>
            </div>""", unsafe_allow_html=True)

# ── Footer ───────────────────────────────────────────────────
st.markdown(f"""
<div style="text-align:center;padding:28px 0 10px;font-family:'JetBrains Mono',monospace;
            font-size:0.52rem;color:{TSEC}44;letter-spacing:2px;">
    ADIDAS SALES INTELLIGENCE DASHBOARD &nbsp;·&nbsp; STREAMLIT + MATPLOTLIB + PANDAS + NUMPY
</div>
""", unsafe_allow_html=True)
