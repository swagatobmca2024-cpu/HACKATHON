import streamlit as st
import pandas as pd
import numpy as np
import io

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sales Data Cleaner",
    page_icon="🧹",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Syne', sans-serif; }

/* Background */
.stApp { background: #0d0d0f; color: #e8e6e0; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: #131316;
    border-right: 1px solid #2a2a30;
}

/* Cards */
.card {
    background: #17171c;
    border: 1px solid #2a2a30;
    border-radius: 12px;
    padding: 20px 24px;
    margin-bottom: 16px;
}
.card-red   { border-left: 4px solid #ff4d4d; }
.card-amber { border-left: 4px solid #ffaa00; }
.card-green { border-left: 4px solid #00cc88; }
.card-blue  { border-left: 4px solid #4d9fff; }

/* Metric pills */
.metric-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 20px; }
.metric-pill {
    background: #1e1e26;
    border: 1px solid #2a2a30;
    border-radius: 8px;
    padding: 12px 20px;
    flex: 1;
    min-width: 140px;
    text-align: center;
}
.metric-pill .val { font-size: 26px; font-weight: 800; }
.metric-pill .lbl { font-size: 11px; color: #888; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
.red    { color: #ff4d4d; }
.amber  { color: #ffaa00; }
.green  { color: #00cc88; }
.blue   { color: #4d9fff; }

/* Section titles */
.section-title {
    font-size: 13px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 2px;
    color: #666;
    margin: 28px 0 12px 0;
    border-bottom: 1px solid #2a2a30;
    padding-bottom: 8px;
}

/* Badges */
.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-family: 'DM Mono', monospace;
    font-weight: 500;
}
.badge-red   { background: #2a0d0d; color: #ff4d4d; border: 1px solid #ff4d4d40; }
.badge-green { background: #0d2a1e; color: #00cc88; border: 1px solid #00cc8840; }
.badge-amber { background: #2a1d0d; color: #ffaa00; border: 1px solid #ffaa0040; }
.badge-blue  { background: #0d1a2a; color: #4d9fff; border: 1px solid #4d9fff40; }

/* Buttons */
.stButton > button {
    background: #00cc88 !important;
    color: #0d0d0f !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 14px !important;
    padding: 10px 24px !important;
    transition: all 0.2s !important;
}
.stButton > button:hover { opacity: 0.85 !important; transform: translateY(-1px) !important; }

/* Dataframe */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* Checkbox */
.stCheckbox label { font-size: 14px !important; }

/* Expander */
.streamlit-expanderHeader { font-size: 14px !important; font-weight: 600 !important; }

/* Success / warning / error boxes */
.stSuccess, .stWarning, .stError, .stInfo { border-radius: 8px !important; }

/* Download button */
.stDownloadButton > button {
    background: #17171c !important;
    color: #00cc88 !important;
    border: 1px solid #00cc88 !important;
    border-radius: 8px !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
}
</style>
""", unsafe_allow_html=True)


# ── Constants ────────────────────────────────────────────────────────────────
VALID_RETAILERS = ['Foot Locker', 'Walmart', 'Sports Direct', 'West Gear', "Kohl's", 'Amazon']
VALID_REGIONS   = ['Northeast', 'South', 'West', 'Midwest', 'Southeast']
VALID_METHODS   = ['In-store', 'Outlet', 'Online']
VALID_PRODUCTS  = [
    "Men's Street Footwear", "Men's Athletic Footwear",
    "Women's Street Footwear", "Women's Athletic Footwear",
    "Men's Apparel", "Women's Apparel"
]
DATE_MIN = pd.Timestamp('2020-01-01')
DATE_MAX = pd.Timestamp('2021-12-31')


# ── Helpers ──────────────────────────────────────────────────────────────────
@st.cache_data
def load_data(file):
    return pd.read_excel(file)


def detect_issues(df: pd.DataFrame) -> dict:
    issues = {}

    # 1. Missing values
    mv = df.isnull().sum()
    issues['missing'] = mv[mv > 0].to_dict()

    # 2. Duplicates
    issues['duplicates'] = int(df.duplicated().sum())

    # 3. Zero / negative units
    issues['zero_units']     = df[df['Units Sold'] <= 0].index.tolist()
    issues['negative_price'] = df[df['Price per Unit'] < 0].index.tolist()
    issues['negative_sales'] = df[df['Total Sales'] < 0].index.tolist()

    # 4. Total Sales mismatch  (Price × Units ≠ Total Sales)
    df2 = df.copy()
    df2['_calc'] = df2['Price per Unit'] * df2['Units Sold']
    issues['sales_mismatch'] = df2[df2['Total Sales'] != df2['_calc']].index.tolist()

    # 5. Operating Margin out of range (0–1)
    issues['margin_oob'] = df[
        (df['Operating Margin'] < 0) | (df['Operating Margin'] > 1)
    ].index.tolist()

    # 6. Invalid categories
    issues['bad_retailer'] = df[~df['Retailer'].isin(VALID_RETAILERS)].index.tolist()
    issues['bad_region']   = df[~df['Region'].isin(VALID_REGIONS)].index.tolist()
    issues['bad_method']   = df[~df['Sales Method'].isin(VALID_METHODS)].index.tolist()
    issues['bad_product']  = df[~df['Product'].isin(VALID_PRODUCTS)].index.tolist()

    # 7. Date out of range
    issues['date_oob'] = df[
        (df['Invoice Date'] < DATE_MIN) | (df['Invoice Date'] > DATE_MAX)
    ].index.tolist()

    # 8. Whitespace in text columns
    ws_cols = ['Retailer', 'Region', 'State', 'City', 'Product', 'Sales Method']
    ws_rows = set()
    for col in ws_cols:
        ws_rows.update(df[df[col] != df[col].str.strip()].index.tolist())
    issues['whitespace'] = list(ws_rows)

    return issues


def total_flagged(issues: dict) -> int:
    flagged = set()
    for key, val in issues.items():
        if key == 'missing':
            continue  # column-level, not row-level
        if isinstance(val, list):
            flagged.update(val)
        elif isinstance(val, int) and key == 'duplicates':
            pass
    return len(flagged)


def apply_fixes(df: pd.DataFrame, choices: dict) -> pd.DataFrame:
    df = df.copy()

    if choices.get('fix_whitespace'):
        for col in ['Retailer', 'Region', 'State', 'City', 'Product', 'Sales Method']:
            df[col] = df[col].str.strip()

    if choices.get('fix_duplicates'):
        df = df.drop_duplicates()

    if choices.get('drop_zero_units'):
        df = df[df['Units Sold'] > 0]

    if choices.get('fix_sales_mismatch'):
        df['Total Sales'] = df['Price per Unit'] * df['Units Sold']

    if choices.get('fix_operating_profit'):
        df['Operating Profit'] = df['Total Sales'] * df['Operating Margin']

    if choices.get('drop_margin_oob'):
        df = df[(df['Operating Margin'] >= 0) & (df['Operating Margin'] <= 1)]

    if choices.get('drop_date_oob'):
        df = df[(df['Invoice Date'] >= DATE_MIN) & (df['Invoice Date'] <= DATE_MAX)]

    if choices.get('drop_bad_categories'):
        df = df[
            df['Retailer'].isin(VALID_RETAILERS) &
            df['Region'].isin(VALID_REGIONS) &
            df['Sales Method'].isin(VALID_METHODS) &
            df['Product'].isin(VALID_PRODUCTS)
        ]

    return df


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Cleaned Data')
    return buf.getvalue()


# ── App ──────────────────────────────────────────────────────────────────────
st.markdown("## 🧹 Sales Data Cleaner")
st.markdown('<p style="color:#666;font-size:14px;margin-top:-12px;">Detect, inspect & fix data quality issues in your sales Excel file</p>', unsafe_allow_html=True)

# ── Sidebar: Upload ───────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📂 Upload File")
    uploaded = st.file_uploader("Drop your Excel file here", type=["xlsx", "xls"])

    st.markdown("---")
    st.markdown("### ℹ️ What this app checks")
    checks = [
        ("🔴", "Zero / negative units"),
        ("🔴", "Total Sales ≠ Price × Units"),
        ("🟡", "Duplicate rows"),
        ("🟡", "Operating margin out of 0–100%"),
        ("🟡", "Dates outside 2020–2021"),
        ("🟢", "Whitespace in text fields"),
        ("🟢", "Invalid category values"),
        ("🟢", "Missing values"),
    ]
    for icon, label in checks:
        st.markdown(f"{icon} {label}")

if not uploaded:
    st.markdown("""
    <div class="card card-blue" style="margin-top:40px;text-align:center;padding:60px;">
        <div style="font-size:48px;margin-bottom:16px;">📊</div>
        <div style="font-size:20px;font-weight:700;margin-bottom:8px;">Upload your Excel file to begin</div>
        <div style="color:#666;font-size:14px;">Supports .xlsx and .xls · Use the sidebar to upload</div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ── Load & Detect ─────────────────────────────────────────────────────────────
df_raw = load_data(uploaded)
issues = detect_issues(df_raw)

n_rows      = len(df_raw)
n_flagged   = total_flagged(issues)
n_missing   = sum(issues['missing'].values()) if issues['missing'] else 0
n_dupes     = issues['duplicates']
n_zero      = len(issues['zero_units'])
n_mismatch  = len(issues['sales_mismatch'])
n_margin    = len(issues['margin_oob'])
n_date      = len(issues['date_oob'])
n_ws        = len(issues['whitespace'])
n_badcat    = len(issues['bad_retailer']) + len(issues['bad_region']) + len(issues['bad_method']) + len(issues['bad_product'])

# ── KPI row ───────────────────────────────────────────────────────────────────
col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Total Rows", f"{n_rows:,}")
with col2:
    st.metric("Flagged Rows", f"{n_flagged:,}", delta=f"{n_flagged/n_rows:.1%} of total", delta_color="inverse")
with col3:
    st.metric("Missing Values", f"{n_missing:,}", delta_color="inverse")
with col4:
    st.metric("Duplicates", f"{n_dupes:,}", delta_color="inverse")
with col5:
    st.metric("Data Errors", f"{n_zero + n_mismatch + n_margin + n_date:,}", delta_color="inverse")

st.markdown("---")

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["🔍 Issue Report", "🛠️ Clean Data", "✅ Preview & Export"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — ISSUE REPORT
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown('<div class="section-title">Critical Issues</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        severity = "🔴 Found" if n_zero > 0 else "✅ None"
        color    = "card-red" if n_zero > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Zero / Negative Units Sold</b><br>
            <span style="font-size:13px;color:#888;">Rows where Units Sold ≤ 0 — no sale occurred</span><br><br>
            <span class="badge {'badge-red' if n_zero > 0 else 'badge-green'}">{severity} · {n_zero} rows</span>
        </div>""", unsafe_allow_html=True)
        if n_zero > 0:
            with st.expander("View affected rows"):
                st.dataframe(df_raw.loc[issues['zero_units']], use_container_width=True)

    with c2:
        severity = "🔴 Found" if n_mismatch > 0 else "✅ None"
        color    = "card-red" if n_mismatch > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Total Sales ≠ Price × Units</b><br>
            <span style="font-size:13px;color:#888;">Calculated revenue doesn't match the stored value</span><br><br>
            <span class="badge {'badge-red' if n_mismatch > 0 else 'badge-green'}">{severity} · {n_mismatch} rows</span>
        </div>""", unsafe_allow_html=True)
        if n_mismatch > 0:
            with st.expander("View affected rows"):
                temp = df_raw.loc[issues['sales_mismatch']].copy()
                temp['Expected Sales'] = temp['Price per Unit'] * temp['Units Sold']
                temp['Difference']     = temp['Total Sales'] - temp['Expected Sales']
                st.dataframe(temp[['Retailer','Product','Price per Unit','Units Sold','Total Sales','Expected Sales','Difference']], use_container_width=True)

    st.markdown('<div class="section-title">Moderate Issues</div>', unsafe_allow_html=True)

    c3, c4, c5 = st.columns(3)

    with c3:
        severity = "🟡 Found" if n_dupes > 0 else "✅ None"
        color    = "card-amber" if n_dupes > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Duplicate Rows</b><br>
            <span style="font-size:13px;color:#888;">Exact row copies that inflate counts</span><br><br>
            <span class="badge {'badge-amber' if n_dupes > 0 else 'badge-green'}">{severity} · {n_dupes} rows</span>
        </div>""", unsafe_allow_html=True)

    with c4:
        severity = "🟡 Found" if n_margin > 0 else "✅ None"
        color    = "card-amber" if n_margin > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Operating Margin Out of Range</b><br>
            <span style="font-size:13px;color:#888;">Margin should be between 0% and 100%</span><br><br>
            <span class="badge {'badge-amber' if n_margin > 0 else 'badge-green'}">{severity} · {n_margin} rows</span>
        </div>""", unsafe_allow_html=True)

    with c5:
        severity = "🟡 Found" if n_date > 0 else "✅ None"
        color    = "card-amber" if n_date > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Dates Outside 2020–2021</b><br>
            <span style="font-size:13px;color:#888;">Invoice dates outside the expected range</span><br><br>
            <span class="badge {'badge-amber' if n_date > 0 else 'badge-green'}">{severity} · {n_date} rows</span>
        </div>""", unsafe_allow_html=True)

    st.markdown('<div class="section-title">Minor Issues</div>', unsafe_allow_html=True)

    c6, c7, c8 = st.columns(3)

    with c6:
        severity = "⚠️ Found" if n_ws > 0 else "✅ None"
        color    = "card-amber" if n_ws > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Whitespace in Text Fields</b><br>
            <span style="font-size:13px;color:#888;">Leading/trailing spaces cause groupby mismatches</span><br><br>
            <span class="badge {'badge-amber' if n_ws > 0 else 'badge-green'}">{severity} · {n_ws} rows</span>
        </div>""", unsafe_allow_html=True)

    with c7:
        severity = "⚠️ Found" if n_badcat > 0 else "✅ None"
        color    = "card-amber" if n_badcat > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Invalid Category Values</b><br>
            <span style="font-size:13px;color:#888;">Retailer / Region / Product / Sales Method anomalies</span><br><br>
            <span class="badge {'badge-amber' if n_badcat > 0 else 'badge-green'}">{severity} · {n_badcat} rows</span>
        </div>""", unsafe_allow_html=True)

    with c8:
        severity = "⚠️ Found" if n_missing > 0 else "✅ None"
        color    = "card-amber" if n_missing > 0 else "card-green"
        st.markdown(f"""
        <div class="card {color}">
            <b>Missing Values</b><br>
            <span style="font-size:13px;color:#888;">Null / NaN cells across all columns</span><br><br>
            <span class="badge {'badge-amber' if n_missing > 0 else 'badge-green'}">{severity} · {n_missing} cells</span>
        </div>""", unsafe_allow_html=True)

        if n_missing > 0:
            with st.expander("See missing breakdown"):
                for col, count in issues['missing'].items():
                    st.markdown(f"**{col}**: {count} missing")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — CLEAN DATA
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("### Choose what to fix")
    st.markdown('<p style="color:#666;font-size:13px;margin-top:-10px;">Select the cleaning operations to apply, then click Apply Fixes.</p>', unsafe_allow_html=True)

    col_a, col_b = st.columns(2)

    with col_a:
        st.markdown('<div class="section-title">🔴 Critical Fixes</div>', unsafe_allow_html=True)

        fix_zero = st.checkbox(
            f"Remove rows with 0 or negative Units Sold  ({n_zero} rows)",
            value=n_zero > 0,
            disabled=n_zero == 0,
        )
        fix_mismatch = st.checkbox(
            f"Recalculate Total Sales = Price × Units  ({n_mismatch} rows affected)",
            value=n_mismatch > 0,
            disabled=n_mismatch == 0,
        )
        if fix_mismatch:
            fix_profit = st.checkbox(
                "Also recalculate Operating Profit = Sales × Margin",
                value=True,
            )
        else:
            fix_profit = False

    with col_b:
        st.markdown('<div class="section-title">🟡 Moderate Fixes</div>', unsafe_allow_html=True)

        fix_dupes = st.checkbox(
            f"Remove duplicate rows  ({n_dupes} rows)",
            value=n_dupes > 0,
            disabled=n_dupes == 0,
        )
        fix_margin = st.checkbox(
            f"Drop rows with margin outside 0–100%  ({n_margin} rows)",
            value=n_margin > 0,
            disabled=n_margin == 0,
        )
        fix_dates = st.checkbox(
            f"Drop rows with dates outside 2020–2021  ({n_date} rows)",
            value=n_date > 0,
            disabled=n_date == 0,
        )

    st.markdown('<div class="section-title">🟢 Minor Fixes</div>', unsafe_allow_html=True)
    col_c, col_d = st.columns(2)

    with col_c:
        fix_ws = st.checkbox(
            f"Strip whitespace from all text columns  ({n_ws} rows)",
            value=n_ws > 0,
            disabled=n_ws == 0,
        )
    with col_d:
        fix_badcat = st.checkbox(
            f"Remove rows with invalid category values  ({n_badcat} rows)",
            value=n_badcat > 0,
            disabled=n_badcat == 0,
        )

    st.markdown("---")

    if st.button("⚡ Apply All Selected Fixes"):
        choices = {
            'drop_zero_units':    fix_zero,
            'fix_sales_mismatch': fix_mismatch,
            'fix_operating_profit': fix_profit,
            'fix_duplicates':     fix_dupes,
            'drop_margin_oob':    fix_margin,
            'drop_date_oob':      fix_dates,
            'fix_whitespace':     fix_ws,
            'drop_bad_categories': fix_badcat,
        }
        df_clean = apply_fixes(df_raw, choices)
        st.session_state['df_clean'] = df_clean
        st.session_state['choices']  = choices

        removed = len(df_raw) - len(df_clean)
        st.success(f"✅ Done! {removed} rows removed. {len(df_clean):,} clean rows remain. Head to **Preview & Export** tab.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — PREVIEW & EXPORT
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    if 'df_clean' not in st.session_state:
        st.info("Apply fixes in the **Clean Data** tab first.")
    else:
        df_clean = st.session_state['df_clean']
        removed  = len(df_raw) - len(df_clean)

        c1, c2, c3 = st.columns(3)
        c1.metric("Original Rows", f"{len(df_raw):,}")
        c2.metric("Rows Removed",  f"{removed:,}", delta=f"-{removed/len(df_raw):.1%}", delta_color="inverse")
        c3.metric("Clean Rows",    f"{len(df_clean):,}")

        st.markdown("### Preview (first 100 rows)")
        st.dataframe(df_clean.head(100), use_container_width=True)

        st.markdown("### Download Cleaned File")
        excel_bytes = to_excel_bytes(df_clean)
        st.download_button(
            label="⬇️ Download cleaned_sales_data.xlsx",
            data=excel_bytes,
            file_name="cleaned_sales_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.markdown("### Column Summary After Cleaning")
        summary = df_clean.describe(include='all').T
        st.dataframe(summary, use_container_width=True)
