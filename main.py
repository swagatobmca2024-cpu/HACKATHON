import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import warnings
warnings.filterwarnings("ignore")

# ─── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Advanced CSV Analytics",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ─── CUSTOM CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    .main { background: #0f1117; }

    .block-container { padding: 1.5rem 2rem 2rem 2rem; max-width: 100%; }

    /* Metric cards */
    .metric-card {
        background: linear-gradient(135deg, #1e2130 0%, #252a3d 100%);
        border: 1px solid #2d3250;
        border-radius: 16px;
        padding: 1.2rem 1.5rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        background: linear-gradient(90deg, #6c63ff, #a78bfa);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        line-height: 1.1;
    }
    .metric-label {
        font-size: 0.75rem;
        font-weight: 500;
        color: #8b92b3;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-top: 0.35rem;
    }
    .metric-delta {
        font-size: 0.78rem;
        color: #34d399;
        margin-top: 0.2rem;
        font-weight: 500;
    }

    /* Section headers */
    .section-header {
        font-size: 1.05rem;
        font-weight: 600;
        color: #c8cde8;
        margin: 1.5rem 0 0.75rem 0;
        padding-bottom: 0.4rem;
        border-bottom: 1px solid #2d3250;
        letter-spacing: 0.03em;
    }

    /* Insight pills */
    .insight-pill {
        display: inline-block;
        background: #1e2130;
        border: 1px solid #3d4275;
        border-radius: 20px;
        padding: 0.25rem 0.75rem;
        font-size: 0.75rem;
        color: #a78bfa;
        margin: 0.2rem;
    }

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 4px;
        background: #1a1e2e;
        border-radius: 12px;
        padding: 4px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        color: #8b92b3;
        font-weight: 500;
        font-size: 0.88rem;
        padding: 0.4rem 1.1rem;
        border: none;
    }
    .stTabs [aria-selected="true"] {
        background: #6c63ff !important;
        color: white !important;
    }

    /* Upload zone */
    .upload-hero {
        background: linear-gradient(135deg, #1a1e2e 0%, #1e2540 100%);
        border: 2px dashed #3d4275;
        border-radius: 20px;
        padding: 3rem;
        text-align: center;
        margin: 2rem 0;
    }
    .upload-title {
        font-size: 1.6rem;
        font-weight: 700;
        color: #e8eaf6;
        margin-bottom: 0.5rem;
    }
    .upload-sub {
        color: #8b92b3;
        font-size: 0.9rem;
    }

    /* Stmetric override */
    [data-testid="stMetricValue"] { font-size: 1.6rem !important; font-weight: 700 !important; color: #a78bfa !important; }
    [data-testid="stMetricLabel"] { font-size: 0.72rem !important; color: #8b92b3 !important; }
    [data-testid="stMetricDelta"] { font-size: 0.78rem !important; }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background: #13161f;
        border-right: 1px solid #1e2130;
    }
    section[data-testid="stSidebar"] .block-container { padding: 1rem; }

    /* Dataframe */
    .stDataFrame { border-radius: 12px; overflow: hidden; }

    /* Alerts / info */
    .stAlert { border-radius: 10px; }

    /* Divider */
    hr { border-color: #1e2130; }

    /* Plotly chart background */
    .js-plotly-plot { border-radius: 14px; overflow: hidden; }
</style>
""", unsafe_allow_html=True)

PLOTLY_THEME = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="#161926",
    font=dict(family="Inter", color="#c8cde8", size=12),
    colorway=["#6c63ff","#f59e0b","#34d399","#f87171","#38bdf8","#fb923c","#a78bfa","#4ade80"],
    margin=dict(l=20, r=20, t=40, b=20),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#c8cde8")),
    xaxis=dict(gridcolor="#1e2130", zerolinecolor="#1e2130", color="#8b92b3"),
    yaxis=dict(gridcolor="#1e2130", zerolinecolor="#1e2130", color="#8b92b3"),
    title=dict(font=dict(size=14, color="#e8eaf6"), x=0.02),
)

def apply_theme(fig, title=""):
    fig.update_layout(**PLOTLY_THEME)
    if title:
        fig.update_layout(title_text=title)
    return fig


# ══════════════════════════════════════════════════════════════════════════════
#  ANALYSIS ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def detect_column_types(df):
    numeric = df.select_dtypes(include=np.number).columns.tolist()
    categorical = df.select_dtypes(include=["object", "category", "bool"]).columns.tolist()
    datetime_cols = []
    for c in df.columns:
        if c not in numeric:
            try:
                parsed = pd.to_datetime(df[c], infer_datetime_format=True, errors="coerce")
                if parsed.notna().sum() / len(df) > 0.5:
                    datetime_cols.append(c)
            except Exception:
                pass
    categorical = [c for c in categorical if c not in datetime_cols]
    return numeric, categorical, datetime_cols


def render_overview(df):
    numeric, categorical, datetime_cols = detect_column_types(df)
    n_rows, n_cols = df.shape
    missing_pct = (df.isnull().sum().sum() / (n_rows * n_cols) * 100)
    duplicates = df.duplicated().sum()

    st.markdown('<div class="section-header">📋 Dataset Overview</div>', unsafe_allow_html=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Rows", f"{n_rows:,}")
    c2.metric("Columns", f"{n_cols}")
    c3.metric("Numeric", f"{len(numeric)}")
    c4.metric("Missing %", f"{missing_pct:.1f}%")
    c5.metric("Duplicates", f"{duplicates:,}")

    col_a, col_b = st.columns([1.6, 1])

    with col_a:
        # Column types bar
        type_counts = {"Numeric": len(numeric), "Categorical": len(categorical), "Datetime": len(datetime_cols)}
        fig = go.Figure(go.Bar(
            x=list(type_counts.keys()),
            y=list(type_counts.values()),
            marker_color=["#6c63ff", "#f59e0b", "#34d399"],
            text=list(type_counts.values()),
            textposition="outside",
        ))
        apply_theme(fig, "Column Type Distribution")
        fig.update_layout(showlegend=False, height=250)
        st.plotly_chart(fig, use_container_width=True)

    with col_b:
        # Missing values heatmap-lite
        miss = df.isnull().sum()
        miss = miss[miss > 0].sort_values(ascending=False)
        if len(miss) > 0:
            fig2 = go.Figure(go.Bar(
                x=miss.values,
                y=miss.index.tolist(),
                orientation="h",
                marker_color="#f87171",
                text=miss.values,
                textposition="outside",
            ))
            apply_theme(fig2, "Missing Values by Column")
            fig2.update_layout(height=250, yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.success("✅ No missing values found!", icon="✅")

    # Data preview
    st.markdown('<div class="section-header">🔍 Data Preview</div>', unsafe_allow_html=True)
    st.dataframe(df.head(20), use_container_width=True, height=300)


def render_numeric_analysis(df):
    numeric, _, _ = detect_column_types(df)
    if not numeric:
        st.info("No numeric columns found.")
        return

    st.markdown('<div class="section-header">📈 Numeric Column Analysis</div>', unsafe_allow_html=True)

    # Descriptive stats table
    desc = df[numeric].describe().T.round(2)
    desc["skewness"] = df[numeric].skew().round(2)
    desc["kurtosis"] = df[numeric].kurtosis().round(2)
    desc["cv%"] = ((df[numeric].std() / df[numeric].mean()) * 100).round(1)
    st.dataframe(desc, use_container_width=True)

    col_sel = st.selectbox("Select column for deep-dive", numeric, key="num_col")

    c1, c2 = st.columns(2)

    with c1:
        fig = px.histogram(df, x=col_sel, nbins=40, title=f"Distribution — {col_sel}",
                           color_discrete_sequence=["#6c63ff"])
        apply_theme(fig)
        fig.update_traces(marker_line_color="#a78bfa", marker_line_width=0.5)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig2 = px.box(df, y=col_sel, title=f"Box Plot — {col_sel}",
                      color_discrete_sequence=["#f59e0b"])
        apply_theme(fig2)
        st.plotly_chart(fig2, use_container_width=True)

    if len(numeric) >= 2:
        st.markdown('<div class="section-header">🔗 Correlation Heatmap</div>', unsafe_allow_html=True)
        corr = df[numeric].corr().round(2)
        fig3 = px.imshow(corr, text_auto=True, aspect="auto",
                         color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
                         title="Pearson Correlation Matrix")
        apply_theme(fig3)
        fig3.update_layout(height=max(350, len(numeric) * 40))
        st.plotly_chart(fig3, use_container_width=True)

        # Scatter
        st.markdown('<div class="section-header">🔵 Scatter Explorer</div>', unsafe_allow_html=True)
        cx, cy = st.columns(2)
        x_col = cx.selectbox("X axis", numeric, key="sx")
        y_col = cy.selectbox("Y axis", numeric, index=min(1, len(numeric)-1), key="sy")
        _, cat_cols, _ = detect_column_types(df)
        color_col = st.selectbox("Color by (optional)", ["None"] + cat_cols, key="sc")
        color = None if color_col == "None" else color_col
        fig4 = px.scatter(df, x=x_col, y=y_col, color=color,
                          trendline="ols", title=f"{x_col} vs {y_col}",
                          opacity=0.7)
        apply_theme(fig4)
        st.plotly_chart(fig4, use_container_width=True)


def render_categorical_analysis(df):
    _, categorical, _ = detect_column_types(df)
    if not categorical:
        st.info("No categorical columns found.")
        return

    st.markdown('<div class="section-header">🏷️ Categorical Column Analysis</div>', unsafe_allow_html=True)

    col_sel = st.selectbox("Select column", categorical, key="cat_col")
    vc = df[col_sel].value_counts().head(25)
    top_pct = (vc.iloc[0] / len(df) * 100) if len(vc) > 0 else 0

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(x=vc.index.astype(str), y=vc.values,
                     labels={"x": col_sel, "y": "Count"},
                     title=f"Value Counts — {col_sel}",
                     color=vc.values,
                     color_continuous_scale="Purples")
        apply_theme(fig)
        fig.update_layout(showlegend=False, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig2 = px.pie(values=vc.values, names=vc.index.astype(str),
                      title=f"Share — {col_sel}",
                      hole=0.45)
        apply_theme(fig2)
        fig2.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown(f"""
    <div style='color:#8b92b3; font-size:0.8rem; margin-top:0.3rem;'>
    <span class='insight-pill'>Unique values: {df[col_sel].nunique()}</span>
    <span class='insight-pill'>Top value: "{vc.index[0]}" ({top_pct:.1f}%)</span>
    <span class='insight-pill'>Null count: {df[col_sel].isnull().sum()}</span>
    </div>
    """, unsafe_allow_html=True)

    # Cross-tab
    if len(categorical) >= 2:
        st.markdown('<div class="section-header">🧩 Cross-Tabulation</div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        row_col = c1.selectbox("Row variable", categorical, key="ct_row")
        col_col = c2.selectbox("Column variable", [c for c in categorical if c != row_col], key="ct_col")
        ct = pd.crosstab(df[row_col], df[col_col])
        fig3 = px.imshow(ct, text_auto=True, aspect="auto",
                         color_continuous_scale="Purples",
                         title=f"{row_col} × {col_col} Cross-tab")
        apply_theme(fig3)
        fig3.update_layout(height=max(300, len(ct) * 35))
        st.plotly_chart(fig3, use_container_width=True)


def render_time_analysis(df):
    _, _, datetime_cols = detect_column_types(df)
    if not datetime_cols:
        st.info("No datetime columns detected.")
        return

    st.markdown('<div class="section-header">🕐 Time Series Analysis</div>', unsafe_allow_html=True)

    dt_col = st.selectbox("Select datetime column", datetime_cols, key="dt_col")
    df2 = df.copy()
    df2[dt_col] = pd.to_datetime(df2[dt_col], errors="coerce")
    df2 = df2.dropna(subset=[dt_col])

    numeric, _, _ = detect_column_types(df2)
    val_col = st.selectbox("Value column (or 'count')", ["— count records —"] + numeric, key="dt_val")

    freq = st.radio("Aggregation", ["D", "W", "ME", "QE"], horizontal=True,
                    format_func=lambda x: {"D":"Daily","W":"Weekly","ME":"Monthly","QE":"Quarterly"}[x])

    df2 = df2.set_index(dt_col)
    if val_col == "— count records —":
        ts = df2.resample(freq).size().reset_index(name="count")
        y = "count"
    else:
        ts = df2[val_col].resample(freq).mean().reset_index()
        y = val_col

    fig = px.line(ts, x=dt_col, y=y, title=f"Trend: {y} over time",
                  markers=True, color_discrete_sequence=["#6c63ff"])
    apply_theme(fig)
    fig.update_traces(line_width=2.5)
    # Add moving average
    if len(ts) >= 5:
        ts["MA"] = ts[y].rolling(min(5, len(ts)//2 or 2), min_periods=1).mean()
        fig.add_scatter(x=ts[dt_col], y=ts["MA"], mode="lines",
                        line=dict(color="#f59e0b", dash="dash", width=2),
                        name="Moving Avg")
    st.plotly_chart(fig, use_container_width=True)


def render_outliers(df):
    numeric, _, _ = detect_column_types(df)
    if not numeric:
        st.info("No numeric columns for outlier detection.")
        return

    st.markdown('<div class="section-header">🚨 Outlier Detection (IQR Method)</div>', unsafe_allow_html=True)

    results = []
    for col in numeric:
        q1, q3 = df[col].quantile(0.25), df[col].quantile(0.75)
        iqr = q3 - q1
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        outs = ((df[col] < lower) | (df[col] > upper)).sum()
        results.append({"Column": col, "Q1": round(q1,2), "Q3": round(q3,2),
                         "IQR": round(iqr,2), "Lower Fence": round(lower,2),
                         "Upper Fence": round(upper,2), "Outliers": int(outs),
                         "Outlier %": round(outs/len(df)*100, 2)})

    res_df = pd.DataFrame(results).sort_values("Outliers", ascending=False)
    st.dataframe(res_df, use_container_width=True)

    if len(numeric) >= 2:
        col_sel = st.selectbox("Visualise outliers for", numeric, key="out_col")
        q1, q3 = df[col_sel].quantile(0.25), df[col_sel].quantile(0.75)
        iqr = q3 - q1
        colors = np.where((df[col_sel] < q1 - 1.5*iqr) | (df[col_sel] > q3 + 1.5*iqr), "#f87171", "#6c63ff")
        fig = go.Figure(go.Scatter(
            x=df.index, y=df[col_sel],
            mode="markers",
            marker=dict(color=colors, size=5, opacity=0.7),
        ))
        apply_theme(fig, f"Outlier Scatter — {col_sel}")
        fig.update_layout(height=300)
        st.plotly_chart(fig, use_container_width=True)


def render_advanced(df):
    numeric, categorical, _ = detect_column_types(df)
    st.markdown('<div class="section-header">🧠 Advanced Analytics</div>', unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs(["📊 Group Aggregation", "📦 Distribution Compare", "🎯 Top N Analysis"])

    with tab1:
        if categorical and numeric:
            gc = st.selectbox("Group by", categorical, key="ga_grp")
            vc = st.multiselect("Aggregate columns", numeric, default=numeric[:min(3,len(numeric))], key="ga_val")
            agg_fn = st.radio("Function", ["mean","sum","count","median","std"], horizontal=True, key="ga_fn")
            if vc:
                grp = df.groupby(gc)[vc].agg(agg_fn).reset_index().round(2)
                st.dataframe(grp, use_container_width=True)
                fig = px.bar(grp, x=gc, y=vc, barmode="group",
                             title=f"{agg_fn.title()} of {', '.join(vc)} by {gc}")
                apply_theme(fig)
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Need at least 1 categorical and 1 numeric column.")

    with tab2:
        if len(numeric) >= 2:
            cols = st.multiselect("Columns to compare", numeric, default=numeric[:min(4,len(numeric))], key="dc_cols")
            if cols:
                fig = go.Figure()
                colors = ["#6c63ff","#f59e0b","#34d399","#f87171","#38bdf8","#fb923c"]
                for i, c in enumerate(cols):
                    fig.add_trace(go.Violin(y=df[c].dropna(), name=c,
                                            box_visible=True, meanline_visible=True,
                                            fillcolor=colors[i % len(colors)],
                                            opacity=0.7, line_color="white"))
                apply_theme(fig, "Violin Distribution Comparison")
                fig.update_layout(height=420, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)

    with tab3:
        if categorical:
            tc = st.selectbox("Column", categorical, key="tn_col")
            n = st.slider("Top N", 5, 30, 10, key="tn_n")
            vc2 = df[tc].value_counts().head(n)
            fig = px.bar(x=vc2.values, y=vc2.index.astype(str),
                         orientation="h", title=f"Top {n} — {tc}",
                         color=vc2.values, color_continuous_scale="Purples",
                         text=vc2.values)
            apply_theme(fig)
            fig.update_layout(showlegend=False, coloraxis_showscale=False,
                               yaxis=dict(autorange="reversed"))
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)


def render_export(df):
    st.markdown('<div class="section-header">💾 Export Processed Data</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        csv = df.to_csv(index=False).encode()
        st.download_button("⬇️ Download CSV", csv, "processed_data.csv", "text/csv", use_container_width=True)

    with c2:
        buf = io.BytesIO()
        df.to_excel(buf, index=False, engine="openpyxl")
        st.download_button("⬇️ Download Excel", buf.getvalue(),
                           "processed_data.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

    with c3:
        summary_buf = io.StringIO()
        numeric, _, _ = detect_column_types(df)
        summary = df[numeric].describe().T if numeric else pd.DataFrame()
        summary.to_csv(summary_buf)
        st.download_button("⬇️ Download Summary Stats", summary_buf.getvalue().encode(),
                           "summary_stats.csv", "text/csv", use_container_width=True)

    st.caption("All exports use your (filtered) dataset as currently loaded.")


# ══════════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 0.5rem 0 1rem 0;'>
        <div style='font-size:2rem;'>📊</div>
        <div style='font-size:1.1rem; font-weight:700; color:#e8eaf6;'>CSV Analytics</div>
        <div style='font-size:0.75rem; color:#6c7293;'>Advanced Data Explorer</div>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader("Upload CSV file", type=["csv", "tsv"],
                                 accept_multiple_files=False)

    if uploaded:
        sep = "\t" if uploaded.name.endswith(".tsv") else ","
        try:
            df_raw = pd.read_csv(uploaded, sep=sep)
            st.success(f"✅ {uploaded.name}")
            st.caption(f"{df_raw.shape[0]:,} rows × {df_raw.shape[1]} cols")

            st.markdown("---")
            st.markdown("**🔧 Filters**")

            numeric, categorical, _ = detect_column_types(df_raw)

            df_filtered = df_raw.copy()
            for cat_col in categorical[:4]:  # up to 4 filter pills
                unique_vals = df_raw[cat_col].dropna().unique().tolist()
                if 2 <= len(unique_vals) <= 30:
                    sel = st.multiselect(cat_col, unique_vals, default=unique_vals, key=f"flt_{cat_col}")
                    df_filtered = df_filtered[df_filtered[cat_col].isin(sel)]

            for num_col in numeric[:2]:
                mn, mx = float(df_raw[num_col].min()), float(df_raw[num_col].max())
                if mn < mx:
                    rng = st.slider(num_col, mn, mx, (mn, mx), key=f"flt_{num_col}")
                    df_filtered = df_filtered[df_filtered[num_col].between(rng[0], rng[1])]

            st.caption(f"Filtered: {len(df_filtered):,} rows")

        except Exception as e:
            st.error(f"Error reading file: {e}")
            df_raw = None
            df_filtered = None
    else:
        df_raw = None
        df_filtered = None

    st.markdown("---")
    st.markdown("""
    <div style='font-size:0.7rem; color:#4a4f6a; text-align:center; padding-top:0.5rem;'>
        Supports CSV · TSV<br>All analysis runs locally
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  MAIN CONTENT
# ══════════════════════════════════════════════════════════════════════════════

if df_filtered is None:
    st.markdown("""
    <div class='upload-hero'>
        <div class='upload-title'>📂 Drop your CSV here</div>
        <div class='upload-sub'>Use the sidebar to upload a CSV or TSV file.<br>
        Get instant charts, stats, correlations, outliers & more.</div>
        <br>
        <div style='display:flex; gap:1rem; justify-content:center; flex-wrap:wrap;'>
            <span class='insight-pill'>📈 Distributions</span>
            <span class='insight-pill'>🔗 Correlations</span>
            <span class='insight-pill'>🕐 Time Series</span>
            <span class='insight-pill'>🚨 Outliers</span>
            <span class='insight-pill'>🧩 Cross-tabs</span>
            <span class='insight-pill'>🎯 Group Agg</span>
            <span class='insight-pill'>💾 Export</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
else:
    df = df_filtered

    tab_ov, tab_num, tab_cat, tab_time, tab_out, tab_adv, tab_exp = st.tabs([
        "📋 Overview",
        "📈 Numeric",
        "🏷️ Categorical",
        "🕐 Time Series",
        "🚨 Outliers",
        "🧠 Advanced",
        "💾 Export"
    ])

    with tab_ov:   render_overview(df)
    with tab_num:  render_numeric_analysis(df)
    with tab_cat:  render_categorical_analysis(df)
    with tab_time: render_time_analysis(df)
    with tab_out:  render_outliers(df)
    with tab_adv:  render_advanced(df)
    with tab_exp:  render_export(df)
