import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import warnings
from io import BytesIO

warnings.filterwarnings('ignore')

st.set_page_config(
    page_title="Enterprise Sales Intelligence Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
body {background:#f4f6f9;}

.metric-card{
background:linear-gradient(135deg, #ffffff 0%, #f8f9fa 100%);
padding:20px;
border-radius:16px;
text-align:center;
box-shadow:0 6px 16px rgba(0,0,0,0.08);
border:1px solid #e5e7eb;
transition:transform 0.2s;
}

.metric-card:hover{
transform:translateY(-2px);
box-shadow:0 8px 20px rgba(0,0,0,0.12);
}

.metric-value{
font-size:34px;
font-weight:800;
color:#1f2937;
margin:8px 0;
}

.metric-label{
font-size:13px;
color:#6b7280;
font-weight:600;
text-transform:uppercase;
letter-spacing:0.5px;
}

.metric-delta{
font-size:13px;
font-weight:600;
margin-top:6px;
}

.metric-delta.positive{color:#10b981;}
.metric-delta.negative{color:#ef4444;}

.section-title{
font-size:28px;
font-weight:800;
margin:24px 0 12px 0;
color:#111827;
border-left:4px solid #3b82f6;
padding-left:12px;
}

.alert-box{
padding:16px;
border-radius:12px;
margin:12px 0;
border-left:4px solid;
font-weight:500;
}

.alert-info{
background:#dbeafe;
border-color:#3b82f6;
color:#1e40af;
}

.alert-warning{
background:#fef3c7;
border-color:#f59e0b;
color:#92400e;
}

.alert-success{
background:#d1fae5;
border-color:#10b981;
color:#065f46;
}

.alert-error{
background:#fee2e2;
border-color:#ef4444;
color:#991b1b;
}

.stat-highlight{
background:#f0f9ff;
padding:12px;
border-radius:8px;
margin:8px 0;
border:1px solid #bfdbfe;
}

.insight-card{
background:#ffffff;
padding:16px;
border-radius:12px;
margin:10px 0;
border-left:3px solid #3b82f6;
box-shadow:0 2px 8px rgba(0,0,0,0.06);
}
</style>
""", unsafe_allow_html=True)

def validate_dataframe(df):
    required_cols = ['Order_Date', 'Customer_Name', 'City', 'Product_Category',
                     'Product_Name', 'Quantity', 'Unit_Price', 'Discount_Percent',
                     'Order_Status', 'Delivery_Days']

    missing_cols = [col for col in required_cols if col not in df.columns]

    if missing_cols:
        return False, f"Missing required columns: {', '.join(missing_cols)}"

    return True, "Data validation successful"

def safe_calculation(func, default=0):
    try:
        result = func()
        if pd.isna(result) or np.isinf(result):
            return default
        return result
    except Exception:
        return default

def calculate_growth_rate(current, previous):
    if previous == 0 or pd.isna(previous) or pd.isna(current):
        return 0
    return ((current - previous) / previous) * 100

def format_currency(value):
    try:
        return f"₹{value:,.0f}"
    except:
        return "₹0"

def format_percentage(value):
    try:
        return f"{value:.1f}%"
    except:
        return "0.0%"

def create_download_button(df, filename="sales_data.csv"):
    csv = df.to_csv(index=False).encode('utf-8')
    return csv

st.markdown("<h1 style='text-align:center; background:linear-gradient(135deg, #667eea 0%, #764ba2 100%); -webkit-background-clip:text; -webkit-text-fill-color:transparent; font-size:42px; font-weight:900;'>🚀 Enterprise Sales Intelligence Dashboard</h1>", unsafe_allow_html=True)

st.markdown("<p style='text-align:center; color:#6b7280; font-size:16px; margin-top:-10px;'>Advanced Analytics & Real-time Insights</p>", unsafe_allow_html=True)

file = st.file_uploader("📤 Upload Excel File (xlsx format)", type=["xlsx"], help="Upload your sales data file")

if file:
    try:
        with st.spinner("Loading and validating data..."):
            df = pd.read_excel(file)

            is_valid, message = validate_dataframe(df)

            if not is_valid:
                st.markdown(f"<div class='alert-box alert-error'>❌ {message}</div>", unsafe_allow_html=True)
                st.stop()

            st.markdown(f"<div class='alert-box alert-success'>✅ Data loaded successfully! {len(df):,} records found.</div>", unsafe_allow_html=True)

        df["Order_Date"] = pd.to_datetime(df["Order_Date"], dayfirst=True, errors="coerce")

        invalid_dates = df["Order_Date"].isna().sum()
        if invalid_dates > 0:
            st.markdown(f"<div class='alert-box alert-warning'>⚠️ Warning: {invalid_dates} records have invalid dates and will be excluded from time-based analysis.</div>", unsafe_allow_html=True)

        df = df.dropna(subset=["Order_Date"])

        df["Clean_Customer_Name"] = (
            df["Customer_Name"].astype(str)
            .str.strip().str.replace(r"\s+", " ", regex=True).str.title()
        )

        df["Quantity"] = pd.to_numeric(df["Quantity"], errors='coerce').fillna(0)
        df["Unit_Price"] = pd.to_numeric(df["Unit_Price"], errors='coerce').fillna(0)
        df["Discount_Percent"] = pd.to_numeric(df["Discount_Percent"], errors='coerce').fillna(0).clip(0, 100)

        df["Total_Amount"] = df["Quantity"] * df["Unit_Price"]
        df["Discount_Amount"] = df["Total_Amount"] * df["Discount_Percent"] / 100
        df["Final_Sales_Amount"] = df["Total_Amount"] - df["Discount_Amount"]

        df["Order_Value_Category"] = np.where(
            df["Final_Sales_Amount"] >= 30000, "High Value",
            np.where(df["Final_Sales_Amount"] >= 10000, "Medium Value", "Low Value")
        )

        df["Delivery_Status"] = np.where(
            df["Order_Status"] == "Cancelled", "Not Delivered",
            np.where(df["Order_Status"] == "Returned", "Returned", "Delivered")
        )

        df["Fast_Delivery_Flag"] = np.where(
            df["Delivery_Days"].isna(), "NA",
            np.where(df["Delivery_Days"] <= 3, "Fast Delivery", "Normal Delivery")
        )

        df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)
        df["Year"] = df["Order_Date"].dt.year
        df["Quarter"] = df["Order_Date"].dt.quarter
        df["Day_of_Week"] = df["Order_Date"].dt.day_name()
        df["Week_of_Year"] = df["Order_Date"].dt.isocalendar().week

        df["Profit_Margin"] = safe_calculation(lambda: (df["Final_Sales_Amount"] / df["Total_Amount"] * 100).fillna(0), 0)

        st.sidebar.header("🔎 Filters")

        date_range = st.sidebar.date_input(
            "Date Range",
            value=(df["Order_Date"].min(), df["Order_Date"].max()),
            min_value=df["Order_Date"].min(),
            max_value=df["Order_Date"].max()
        )

        if len(date_range) == 2:
            df = df[(df["Order_Date"] >= pd.Timestamp(date_range[0])) &
                    (df["Order_Date"] <= pd.Timestamp(date_range[1]))]

        city = st.sidebar.multiselect("City", sorted(df["City"].unique()), help="Select one or more cities")
        category = st.sidebar.multiselect("Category", sorted(df["Product_Category"].unique()), help="Select product categories")
        order_status = st.sidebar.multiselect("Order Status", sorted(df["Order_Status"].unique()), help="Filter by order status")
        value_category = st.sidebar.multiselect("Value Category", ["High Value", "Medium Value", "Low Value"], help="Filter by order value")

        if city:
            df = df[df["City"].isin(city)]
        if category:
            df = df[df["Product_Category"].isin(category)]
        if order_status:
            df = df[df["Order_Status"].isin(order_status)]
        if value_category:
            df = df[df["Order_Value_Category"].isin(value_category)]

        st.sidebar.markdown("---")
        st.sidebar.download_button(
            label="📥 Download Filtered Data",
            data=create_download_button(df),
            file_name="filtered_sales_data.csv",
            mime="text/csv"
        )

        if len(df) == 0:
            st.markdown("<div class='alert-box alert-warning'>⚠️ No data matches your filter criteria. Please adjust filters.</div>", unsafe_allow_html=True)
            st.stop()

        total_sales = safe_calculation(lambda: df["Final_Sales_Amount"].sum(), 0)
        total_orders = len(df)
        total_items = safe_calculation(lambda: df["Quantity"].sum(), 0)
        avg_order = safe_calculation(lambda: df["Final_Sales_Amount"].mean(), 0)
        discount_given = safe_calculation(lambda: df["Discount_Amount"].sum(), 0)
        customers = safe_calculation(lambda: df["Clean_Customer_Name"].nunique(), 0)

        avg_delivery_days = safe_calculation(lambda: df["Delivery_Days"].mean(), 0)
        conversion_rate = safe_calculation(lambda: (df["Order_Status"] == "Delivered").sum() / len(df) * 100, 0)
        return_rate = safe_calculation(lambda: (df["Order_Status"] == "Returned").sum() / len(df) * 100, 0)

        monthly_df = df.groupby("Month").agg({
            "Final_Sales_Amount": "sum",
            "Order_Date": "count"
        }).reset_index()

        if len(monthly_df) >= 2:
            current_month_sales = monthly_df.iloc[-1]["Final_Sales_Amount"]
            previous_month_sales = monthly_df.iloc[-2]["Final_Sales_Amount"]
            sales_growth = calculate_growth_rate(current_month_sales, previous_month_sales)
        else:
            sales_growth = 0

        st.markdown("<div class='section-title'>📌 Key Performance Indicators</div>", unsafe_allow_html=True)

        k1, k2, k3, k4 = st.columns(4)

        k1.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{format_currency(total_sales)}</div>
            <div class='metric-label'>Total Revenue</div>
            <div class='metric-delta {"positive" if sales_growth >= 0 else "negative"}'>
                {"📈" if sales_growth >= 0 else "📉"} {format_percentage(abs(sales_growth))} MoM
            </div>
        </div>
        """, unsafe_allow_html=True)

        k2.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{total_orders:,}</div>
            <div class='metric-label'>Total Orders</div>
            <div class='metric-delta positive'>
                {format_currency(avg_order)} avg
            </div>
        </div>
        """, unsafe_allow_html=True)

        k3.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{customers:,}</div>
            <div class='metric-label'>Unique Customers</div>
            <div class='metric-delta positive'>
                {safe_calculation(lambda: total_orders/customers if customers > 0 else 0, 0):.1f} orders/customer
            </div>
        </div>
        """, unsafe_allow_html=True)

        k4.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{format_percentage(conversion_rate)}</div>
            <div class='metric-label'>Conversion Rate</div>
            <div class='metric-delta {"negative" if return_rate > 5 else "positive"}'>
                {format_percentage(return_rate)} returns
            </div>
        </div>
        """, unsafe_allow_html=True)

        k5, k6, k7, k8 = st.columns(4)

        k5.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{int(total_items):,}</div>
            <div class='metric-label'>Items Sold</div>
            <div class='metric-delta positive'>
                {safe_calculation(lambda: total_items/total_orders if total_orders > 0 else 0, 0):.1f} items/order
            </div>
        </div>
        """, unsafe_allow_html=True)

        k6.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{format_currency(discount_given)}</div>
            <div class='metric-label'>Total Discounts</div>
            <div class='metric-delta {"negative" if discount_given/total_sales*100 > 15 else "positive"}'>
                {format_percentage(safe_calculation(lambda: discount_given/total_sales*100 if total_sales > 0 else 0, 0))} of sales
            </div>
        </div>
        """, unsafe_allow_html=True)

        k7.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{avg_delivery_days:.1f}</div>
            <div class='metric-label'>Avg Delivery Days</div>
            <div class='metric-delta {"positive" if avg_delivery_days <= 3 else "negative"}'>
                {"⚡ Fast" if avg_delivery_days <= 3 else "🐌 Slow"}
            </div>
        </div>
        """, unsafe_allow_html=True)

        k8.markdown(f"""
        <div class='metric-card'>
            <div class='metric-value'>{len(df["Product_Name"].unique()):,}</div>
            <div class='metric-label'>Product Variety</div>
            <div class='metric-delta positive'>
                {len(df["Product_Category"].unique())} categories
            </div>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
            ["📊 Overview", "📈 Trends & Forecasting", "🏆 Products", "👥 Customers", "🚚 Delivery", "🎯 Advanced Analytics"]
        )

        with tab1:
            st.markdown("<div class='section-title'>Sales Performance Overview</div>", unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                ovc = df["Order_Value_Category"].value_counts().reset_index()
                ovc.columns = ["Category", "Count"]
                ovc["Percentage"] = (ovc["Count"] / ovc["Count"].sum() * 100).round(1)

                fig1 = go.Figure(data=[go.Bar(
                    x=ovc["Category"],
                    y=ovc["Count"],
                    text=[f'{count}<br>({pct}%)' for count, pct in zip(ovc["Count"], ovc["Percentage"])],
                    textposition='auto',
                    marker=dict(
                        color=ovc["Count"],
                        colorscale='Blues',
                        showscale=False
                    ),
                    hovertemplate='<b>%{x}</b><br>Orders: %{y}<br>Percentage: %{text}<extra></extra>'
                )])

                fig1.update_layout(
                    title="Orders by Value Category",
                    xaxis_title="Category",
                    yaxis_title="Number of Orders",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig1, use_container_width=True)

            with col2:
                city_sales = df.groupby("City")["Final_Sales_Amount"].sum().reset_index()
                city_sales = city_sales.sort_values("Final_Sales_Amount", ascending=False).head(15)

                fig2 = go.Figure(data=[go.Bar(
                    x=city_sales["City"],
                    y=city_sales["Final_Sales_Amount"],
                    text=[format_currency(val) for val in city_sales["Final_Sales_Amount"]],
                    textposition='auto',
                    marker=dict(
                        color=city_sales["Final_Sales_Amount"],
                        colorscale='Viridis',
                        showscale=True,
                        colorbar=dict(title="Sales (₹)")
                    ),
                    hovertemplate='<b>%{x}</b><br>Sales: ₹%{y:,.0f}<extra></extra>'
                )])

                fig2.update_layout(
                    title="Top 15 Cities by Sales",
                    xaxis_title="City",
                    yaxis_title="Sales Amount (₹)",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig2, use_container_width=True)

            col3, col4 = st.columns(2)

            with col3:
                category_metrics = df.groupby("Product_Category").agg({
                    "Final_Sales_Amount": "sum",
                    "Order_Date": "count"
                }).reset_index()
                category_metrics.columns = ["Category", "Sales", "Orders"]

                fig_tree = px.treemap(
                    category_metrics,
                    path=["Category"],
                    values="Sales",
                    color="Orders",
                    color_continuous_scale="RdYlGn",
                    title="Product Category Distribution (Size: Sales, Color: Orders)"
                )

                fig_tree.update_layout(height=400)
                st.plotly_chart(fig_tree, use_container_width=True)

            with col4:
                status_data = df.groupby("Order_Status").agg({
                    "Final_Sales_Amount": "sum",
                    "Order_Date": "count"
                }).reset_index()
                status_data.columns = ["Status", "Sales", "Orders"]

                fig_status = go.Figure(data=[go.Pie(
                    labels=status_data["Status"],
                    values=status_data["Orders"],
                    hole=0.4,
                    marker=dict(colors=['#10b981', '#ef4444', '#f59e0b', '#3b82f6']),
                    textposition='auto',
                    textinfo='label+percent',
                    hovertemplate='<b>%{label}</b><br>Orders: %{value}<br>Percentage: %{percent}<extra></extra>'
                )])

                fig_status.update_layout(
                    title="Order Status Distribution",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_status, use_container_width=True)

            st.markdown("<div class='section-title'>City vs Category Sales Heatmap</div>", unsafe_allow_html=True)

            heat = df.pivot_table(
                index="City",
                columns="Product_Category",
                values="Final_Sales_Amount",
                aggfunc="sum",
                fill_value=0
            )

            fig_heat = go.Figure(data=go.Heatmap(
                z=heat.values,
                x=heat.columns,
                y=heat.index,
                colorscale='YlOrRd',
                text=heat.values,
                texttemplate='₹%{text:,.0f}',
                textfont={"size": 10},
                hovertemplate='City: %{y}<br>Category: %{x}<br>Sales: ₹%{z:,.0f}<extra></extra>'
            ))

            fig_heat.update_layout(
                title="Sales Heatmap: City vs Product Category",
                xaxis_title="Product Category",
                yaxis_title="City",
                height=600,
                template="plotly_white"
            )

            st.plotly_chart(fig_heat, use_container_width=True)

        with tab2:
            st.markdown("<div class='section-title'>Sales Trends & Performance</div>", unsafe_allow_html=True)

            monthly = df.groupby("Month").agg({
                "Final_Sales_Amount": "sum",
                "Order_Date": "count",
                "Quantity": "sum"
            }).reset_index()
            monthly.columns = ["Month", "Sales", "Orders", "Items"]

            if len(monthly) >= 3:
                monthly["Sales_MA"] = monthly["Sales"].rolling(window=3, min_periods=1).mean()

            fig3 = make_subplots(
                rows=2, cols=1,
                subplot_titles=("Monthly Revenue Trend", "Order Volume Trend"),
                vertical_spacing=0.15,
                specs=[[{"secondary_y": True}], [{"secondary_y": False}]]
            )

            fig3.add_trace(
                go.Scatter(
                    x=monthly["Month"],
                    y=monthly["Sales"],
                    name="Sales",
                    mode='lines+markers',
                    line=dict(color='#3b82f6', width=3),
                    marker=dict(size=8),
                    hovertemplate='Month: %{x}<br>Sales: ₹%{y:,.0f}<extra></extra>'
                ),
                row=1, col=1, secondary_y=False
            )

            if len(monthly) >= 3:
                fig3.add_trace(
                    go.Scatter(
                        x=monthly["Month"],
                        y=monthly["Sales_MA"],
                        name="3-Month Moving Average",
                        mode='lines',
                        line=dict(color='#10b981', width=2, dash='dash'),
                        hovertemplate='Month: %{x}<br>MA: ₹%{y:,.0f}<extra></extra>'
                    ),
                    row=1, col=1, secondary_y=False
                )

            fig3.add_trace(
                go.Bar(
                    x=monthly["Month"],
                    y=monthly["Orders"],
                    name="Orders",
                    marker=dict(color='#8b5cf6'),
                    hovertemplate='Month: %{x}<br>Orders: %{y}<extra></extra>'
                ),
                row=2, col=1
            )

            fig3.update_layout(
                height=700,
                template="plotly_white",
                showlegend=True,
                hovermode='x unified'
            )

            st.plotly_chart(fig3, use_container_width=True)

            col1, col2 = st.columns(2)

            with col1:
                day_of_week_sales = df.groupby("Day_of_Week")["Final_Sales_Amount"].sum().reset_index()
                day_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                day_of_week_sales["Day_of_Week"] = pd.Categorical(day_of_week_sales["Day_of_Week"], categories=day_order, ordered=True)
                day_of_week_sales = day_of_week_sales.sort_values("Day_of_Week")

                fig_dow = go.Figure(data=[go.Bar(
                    x=day_of_week_sales["Day_of_Week"],
                    y=day_of_week_sales["Final_Sales_Amount"],
                    marker=dict(
                        color=day_of_week_sales["Final_Sales_Amount"],
                        colorscale='Sunset',
                        showscale=False
                    ),
                    text=[format_currency(val) for val in day_of_week_sales["Final_Sales_Amount"]],
                    textposition='auto',
                    hovertemplate='<b>%{x}</b><br>Sales: ₹%{y:,.0f}<extra></extra>'
                )])

                fig_dow.update_layout(
                    title="Sales by Day of Week",
                    xaxis_title="Day",
                    yaxis_title="Sales (₹)",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_dow, use_container_width=True)

            with col2:
                quarterly_sales = df.groupby(["Year", "Quarter"]).agg({
                    "Final_Sales_Amount": "sum"
                }).reset_index()
                quarterly_sales["Period"] = quarterly_sales["Year"].astype(str) + " Q" + quarterly_sales["Quarter"].astype(str)

                fig_quarter = go.Figure(data=[go.Bar(
                    x=quarterly_sales["Period"],
                    y=quarterly_sales["Final_Sales_Amount"],
                    marker=dict(
                        color=quarterly_sales["Final_Sales_Amount"],
                        colorscale='Teal',
                        showscale=False
                    ),
                    text=[format_currency(val) for val in quarterly_sales["Final_Sales_Amount"]],
                    textposition='auto',
                    hovertemplate='<b>%{x}</b><br>Sales: ₹%{y:,.0f}<extra></extra>'
                )])

                fig_quarter.update_layout(
                    title="Quarterly Sales Performance",
                    xaxis_title="Quarter",
                    yaxis_title="Sales (₹)",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_quarter, use_container_width=True)

            st.markdown("<div class='insight-card'>💡 <b>Insight:</b> Use these trends to identify seasonal patterns and plan inventory accordingly.</div>", unsafe_allow_html=True)

        with tab3:
            st.markdown("<div class='section-title'>Product Performance Analysis</div>", unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                prod = df.groupby("Product_Name").agg({
                    "Final_Sales_Amount": "sum",
                    "Quantity": "sum",
                    "Order_Date": "count"
                }).reset_index()
                prod.columns = ["Product", "Sales", "Quantity", "Orders"]
                prod = prod.sort_values(by="Sales", ascending=False).head(15)

                fig4 = go.Figure(data=[go.Bar(
                    y=prod["Product"],
                    x=prod["Sales"],
                    orientation='h',
                    marker=dict(
                        color=prod["Sales"],
                        colorscale='Blues',
                        showscale=True,
                        colorbar=dict(title="Sales (₹)")
                    ),
                    text=[format_currency(val) for val in prod["Sales"]],
                    textposition='auto',
                    hovertemplate='<b>%{y}</b><br>Sales: ₹%{x:,.0f}<br>Qty: %{customdata[0]}<br>Orders: %{customdata[1]}<extra></extra>',
                    customdata=prod[["Quantity", "Orders"]].values
                )])

                fig4.update_layout(
                    title="Top 15 Products by Revenue",
                    xaxis_title="Sales (₹)",
                    yaxis_title="Product",
                    height=500,
                    template="plotly_white"
                )

                st.plotly_chart(fig4, use_container_width=True)

            with col2:
                prod_qty = prod.sort_values("Quantity", ascending=False).head(15)

                fig_qty = go.Figure(data=[go.Bar(
                    y=prod_qty["Product"],
                    x=prod_qty["Quantity"],
                    orientation='h',
                    marker=dict(
                        color=prod_qty["Quantity"],
                        colorscale='Greens',
                        showscale=True,
                        colorbar=dict(title="Quantity")
                    ),
                    text=prod_qty["Quantity"],
                    textposition='auto',
                    hovertemplate='<b>%{y}</b><br>Quantity Sold: %{x:,.0f}<extra></extra>'
                )])

                fig_qty.update_layout(
                    title="Top 15 Products by Quantity Sold",
                    xaxis_title="Quantity",
                    yaxis_title="Product",
                    height=500,
                    template="plotly_white"
                )

                st.plotly_chart(fig_qty, use_container_width=True)

            category_performance = df.groupby("Product_Category").agg({
                "Final_Sales_Amount": ["sum", "mean"],
                "Quantity": "sum",
                "Order_Date": "count"
            }).reset_index()
            category_performance.columns = ["Category", "Total Sales", "Avg Order Value", "Total Quantity", "Orders"]

            fig_sunburst = px.sunburst(
                df,
                path=["Product_Category", "Product_Name"],
                values="Final_Sales_Amount",
                color="Final_Sales_Amount",
                color_continuous_scale="RdYlGn",
                title="Product Hierarchy: Category > Product (Size = Sales)"
            )

            fig_sunburst.update_layout(height=600)
            st.plotly_chart(fig_sunburst, use_container_width=True)

            st.markdown("<div class='section-title'>Category Performance Table</div>", unsafe_allow_html=True)

            category_performance["Total Sales"] = category_performance["Total Sales"].apply(format_currency)
            category_performance["Avg Order Value"] = category_performance["Avg Order Value"].apply(format_currency)

            st.dataframe(
                category_performance.style.set_properties(**{
                    'background-color': '#f8f9fa',
                    'color': '#1f2937',
                    'border': '1px solid #e5e7eb'
                }),
                use_container_width=True
            )

        with tab4:
            st.markdown("<div class='section-title'>Customer Intelligence</div>", unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                cust = df.groupby("Clean_Customer_Name").agg({
                    "Final_Sales_Amount": "sum",
                    "Order_Date": "count",
                    "Quantity": "sum"
                }).reset_index()
                cust.columns = ["Customer", "Total Sales", "Orders", "Items"]
                cust = cust.sort_values(by="Total Sales", ascending=False).head(15)

                fig5 = go.Figure(data=[go.Bar(
                    y=cust["Customer"],
                    x=cust["Total Sales"],
                    orientation='h',
                    marker=dict(
                        color=cust["Total Sales"],
                        colorscale='Purples',
                        showscale=True,
                        colorbar=dict(title="Sales (₹)")
                    ),
                    text=[format_currency(val) for val in cust["Total Sales"]],
                    textposition='auto',
                    hovertemplate='<b>%{y}</b><br>Sales: ₹%{x:,.0f}<br>Orders: %{customdata[0]}<br>Items: %{customdata[1]}<extra></extra>',
                    customdata=cust[["Orders", "Items"]].values
                )])

                fig5.update_layout(
                    title="Top 15 Customers by Revenue",
                    xaxis_title="Total Sales (₹)",
                    yaxis_title="Customer",
                    height=500,
                    template="plotly_white"
                )

                st.plotly_chart(fig5, use_container_width=True)

            with col2:
                cust_freq = cust.sort_values("Orders", ascending=False).head(15)

                fig_freq = go.Figure(data=[go.Bar(
                    y=cust_freq["Customer"],
                    x=cust_freq["Orders"],
                    orientation='h',
                    marker=dict(
                        color=cust_freq["Orders"],
                        colorscale='Oranges',
                        showscale=True,
                        colorbar=dict(title="Orders")
                    ),
                    text=cust_freq["Orders"],
                    textposition='auto',
                    hovertemplate='<b>%{y}</b><br>Orders: %{x}<extra></extra>'
                )])

                fig_freq.update_layout(
                    title="Most Frequent Customers (by Order Count)",
                    xaxis_title="Number of Orders",
                    yaxis_title="Customer",
                    height=500,
                    template="plotly_white"
                )

                st.plotly_chart(fig_freq, use_container_width=True)

            customer_segmentation = df.groupby("Clean_Customer_Name").agg({
                "Final_Sales_Amount": "sum",
                "Order_Date": "count"
            }).reset_index()
            customer_segmentation.columns = ["Customer", "Total_Sales", "Order_Count"]

            customer_segmentation["Avg_Order_Value"] = customer_segmentation["Total_Sales"] / customer_segmentation["Order_Count"]

            fig_segment = px.scatter(
                customer_segmentation,
                x="Order_Count",
                y="Total_Sales",
                size="Avg_Order_Value",
                color="Order_Count",
                hover_name="Customer",
                color_continuous_scale="Viridis",
                title="Customer Segmentation (Size = Avg Order Value)",
                labels={"Order_Count": "Number of Orders", "Total_Sales": "Total Sales (₹)"}
            )

            fig_segment.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig_segment, use_container_width=True)

            st.markdown("<div class='insight-card'>💡 <b>Customer Insight:</b> Focus on high-value, frequent customers for loyalty programs and personalized offers.</div>", unsafe_allow_html=True)

        with tab5:
            st.markdown("<div class='section-title'>Delivery & Operations</div>", unsafe_allow_html=True)

            col1, col2 = st.columns(2)

            with col1:
                fast = df["Fast_Delivery_Flag"].value_counts().reset_index()
                fast.columns = ["Type", "Count"]

                fig6 = go.Figure(data=[go.Pie(
                    labels=fast["Type"],
                    values=fast["Count"],
                    hole=0.5,
                    marker=dict(colors=['#10b981', '#f59e0b', '#6b7280']),
                    textposition='auto',
                    textinfo='label+percent+value',
                    hovertemplate='<b>%{label}</b><br>Orders: %{value}<br>Percentage: %{percent}<extra></extra>'
                )])

                fig6.update_layout(
                    title="Delivery Speed Distribution",
                    height=400,
                    template="plotly_white",
                    annotations=[dict(text='Delivery<br>Speed', x=0.5, y=0.5, font_size=16, showarrow=False)]
                )

                st.plotly_chart(fig6, use_container_width=True)

            with col2:
                delv = df["Delivery_Status"].value_counts().reset_index()
                delv.columns = ["Status", "Count"]

                fig7 = go.Figure(data=[go.Bar(
                    x=delv["Status"],
                    y=delv["Count"],
                    marker=dict(
                        color=['#10b981', '#ef4444', '#f59e0b'],
                    ),
                    text=delv["Count"],
                    textposition='auto',
                    hovertemplate='<b>%{x}</b><br>Orders: %{y}<extra></extra>'
                )])

                fig7.update_layout(
                    title="Delivery Status Overview",
                    xaxis_title="Status",
                    yaxis_title="Number of Orders",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig7, use_container_width=True)

            delivery_by_city = df.groupby("City").agg({
                "Delivery_Days": "mean"
            }).reset_index()
            delivery_by_city.columns = ["City", "Avg_Delivery_Days"]
            delivery_by_city = delivery_by_city.sort_values("Avg_Delivery_Days", ascending=True).head(20)

            fig_city_delivery = go.Figure(data=[go.Bar(
                x=delivery_by_city["City"],
                y=delivery_by_city["Avg_Delivery_Days"],
                marker=dict(
                    color=delivery_by_city["Avg_Delivery_Days"],
                    colorscale='RdYlGn_r',
                    showscale=True,
                    colorbar=dict(title="Days")
                ),
                text=[f'{val:.1f}' for val in delivery_by_city["Avg_Delivery_Days"]],
                textposition='auto',
                hovertemplate='<b>%{x}</b><br>Avg Days: %{y:.1f}<extra></extra>'
            )])

            fig_city_delivery.update_layout(
                title="Average Delivery Time by City (Top 20)",
                xaxis_title="City",
                yaxis_title="Average Days",
                height=400,
                template="plotly_white"
            )

            st.plotly_chart(fig_city_delivery, use_container_width=True)

            delivery_performance = df.groupby("Order_Value_Category").agg({
                "Delivery_Days": ["mean", "median", "min", "max"]
            }).reset_index()
            delivery_performance.columns = ["Category", "Mean", "Median", "Min", "Max"]

            st.markdown("<div class='section-title'>Delivery Performance by Order Value</div>", unsafe_allow_html=True)
            st.dataframe(
                delivery_performance.style.format({
                    "Mean": "{:.1f}",
                    "Median": "{:.1f}",
                    "Min": "{:.0f}",
                    "Max": "{:.0f}"
                }).set_properties(**{
                    'background-color': '#f8f9fa',
                    'color': '#1f2937',
                    'border': '1px solid #e5e7eb'
                }),
                use_container_width=True
            )

        with tab6:
            try:
                st.markdown("<div class='section-title'>Advanced Analytics & Insights</div>", unsafe_allow_html=True)

                col1, col2 = st.columns(2)

                with col1:
                    discount_data_clean = df[df["Discount_Percent"].notna()].copy()

                    if len(discount_data_clean) > 0:
                        max_discount = discount_data_clean["Discount_Percent"].max()

                        if max_discount <= 20:
                            bins = [0, 5, 10, 15, max_discount + 0.01]
                            labels = ["0-5%", "5-10%", "10-15%", f"15-{max_discount:.0f}%"]
                        else:
                            bins = [0, 5, 10, 15, 20, max_discount + 0.01]
                            labels = ["0-5%", "5-10%", "10-15%", "15-20%", f"20-{max_discount:.0f}%"]

                        discount_data_clean["Discount_Range"] = pd.cut(
                            discount_data_clean["Discount_Percent"],
                            bins=bins,
                            labels=labels,
                            include_lowest=True,
                            duplicates='drop'
                        )

                        discount_impact = discount_data_clean.groupby("Discount_Range", observed=True).agg({
                            "Final_Sales_Amount": "sum",
                            "Order_Date": "count"
                        }).reset_index()
                        discount_impact.columns = ["Discount Range", "Sales", "Orders"]

                        fig_discount = go.Figure()

                        fig_discount.add_trace(go.Bar(
                            x=discount_impact["Discount Range"],
                            y=discount_impact["Sales"],
                            name="Sales",
                            marker_color='#3b82f6',
                            yaxis='y',
                            hovertemplate='Range: %{x}<br>Sales: ₹%{y:,.0f}<extra></extra>'
                        ))

                        fig_discount.add_trace(go.Scatter(
                            x=discount_impact["Discount Range"],
                            y=discount_impact["Orders"],
                            name="Orders",
                            mode='lines+markers',
                            marker_color='#ef4444',
                            yaxis='y2',
                            hovertemplate='Range: %{x}<br>Orders: %{y}<extra></extra>'
                        ))

                        fig_discount.update_layout(
                            title="Discount Impact Analysis",
                            xaxis_title="Discount Range (%)",
                            yaxis=dict(title="Sales (₹)", side='left'),
                            yaxis2=dict(title="Orders", overlaying='y', side='right'),
                            height=400,
                            template="plotly_white",
                            hovermode='x unified'
                        )

                        st.plotly_chart(fig_discount, use_container_width=True)
                    else:
                        st.info("No discount data available for analysis")

            with col2:
                basket_size = df.groupby("Order_Date").agg({
                    "Quantity": "sum"
                }).reset_index()
                basket_size = basket_size.groupby("Quantity").size().reset_index()
                basket_size.columns = ["Items per Order", "Frequency"]
                basket_size = basket_size[basket_size["Items per Order"] <= 20]

                fig_basket = go.Figure(data=[go.Bar(
                    x=basket_size["Items per Order"],
                    y=basket_size["Frequency"],
                    marker=dict(
                        color=basket_size["Frequency"],
                        colorscale='Mint',
                        showscale=False
                    ),
                    hovertemplate='Items: %{x}<br>Frequency: %{y}<extra></extra>'
                )])

                fig_basket.update_layout(
                    title="Basket Size Distribution",
                    xaxis_title="Items per Order",
                    yaxis_title="Frequency",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_basket, use_container_width=True)

            sales_funnel = pd.DataFrame({
                "Stage": ["Total Orders", "Delivered", "High Value", "Repeat Customers"],
                "Count": [
                    total_orders,
                    (df["Order_Status"] == "Delivered").sum(),
                    (df["Order_Value_Category"] == "High Value").sum(),
                    len(df[df.groupby("Clean_Customer_Name")["Order_Date"].transform("count") > 1])
                ]
            })

            fig_funnel = go.Figure(go.Funnel(
                y=sales_funnel["Stage"],
                x=sales_funnel["Count"],
                textposition="inside",
                textinfo="value+percent initial",
                marker=dict(color=["#3b82f6", "#10b981", "#f59e0b", "#8b5cf6"]),
                hovertemplate='<b>%{y}</b><br>Count: %{x}<br>%{percentInitial}<extra></extra>'
            ))

            fig_funnel.update_layout(
                title="Sales Funnel Analysis",
                height=500,
                template="plotly_white"
            )

            st.plotly_chart(fig_funnel, use_container_width=True)

            col3, col4 = st.columns(2)

            with col3:
                rfm_data = df.groupby("Clean_Customer_Name").agg({
                    "Order_Date": lambda x: (df["Order_Date"].max() - x.max()).days,
                    "Final_Sales_Amount": ["sum", "count"]
                }).reset_index()
                rfm_data.columns = ["Customer", "Recency", "Monetary", "Frequency"]

                rfm_data["R_Score"] = pd.qcut(rfm_data["Recency"], q=3, labels=[3, 2, 1], duplicates='drop')
                rfm_data["F_Score"] = pd.qcut(rfm_data["Frequency"], q=3, labels=[1, 2, 3], duplicates='drop')
                rfm_data["M_Score"] = pd.qcut(rfm_data["Monetary"], q=3, labels=[1, 2, 3], duplicates='drop')

                rfm_data["RFM_Score"] = (rfm_data["R_Score"].astype(int) +
                                         rfm_data["F_Score"].astype(int) +
                                         rfm_data["M_Score"].astype(int))

                rfm_segments = rfm_data["RFM_Score"].value_counts().sort_index().reset_index()
                rfm_segments.columns = ["RFM Score", "Customers"]

                fig_rfm = go.Figure(data=[go.Bar(
                    x=rfm_segments["RFM Score"],
                    y=rfm_segments["Customers"],
                    marker=dict(
                        color=rfm_segments["RFM Score"],
                        colorscale='Tealgrn',
                        showscale=False
                    ),
                    text=rfm_segments["Customers"],
                    textposition='auto',
                    hovertemplate='RFM Score: %{x}<br>Customers: %{y}<extra></extra>'
                )])

                fig_rfm.update_layout(
                    title="RFM Score Distribution (Recency, Frequency, Monetary)",
                    xaxis_title="RFM Score (3=Low, 9=High)",
                    yaxis_title="Number of Customers",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_rfm, use_container_width=True)

            with col4:
                monthly_growth = monthly.copy()
                monthly_growth["Growth_Rate"] = monthly_growth["Sales"].pct_change() * 100
                monthly_growth = monthly_growth.dropna()

                fig_growth = go.Figure()

                fig_growth.add_trace(go.Bar(
                    x=monthly_growth["Month"],
                    y=monthly_growth["Growth_Rate"],
                    marker=dict(
                        color=monthly_growth["Growth_Rate"],
                        colorscale='RdYlGn',
                        cmid=0,
                        showscale=True,
                        colorbar=dict(title="Growth %")
                    ),
                    text=[f'{val:.1f}%' for val in monthly_growth["Growth_Rate"]],
                    textposition='auto',
                    hovertemplate='Month: %{x}<br>Growth: %{y:.1f}%<extra></extra>'
                ))

                fig_growth.update_layout(
                    title="Month-over-Month Growth Rate",
                    xaxis_title="Month",
                    yaxis_title="Growth Rate (%)",
                    height=400,
                    template="plotly_white"
                )

                st.plotly_chart(fig_growth, use_container_width=True)

            st.markdown("<div class='section-title'>Key Business Insights</div>", unsafe_allow_html=True)

            insights_col1, insights_col2, insights_col3 = st.columns(3)

            with insights_col1:
                top_performing_city = df.groupby("City")["Final_Sales_Amount"].sum().idxmax()
                top_city_sales = df.groupby("City")["Final_Sales_Amount"].sum().max()

                st.markdown(f"""
                <div class='insight-card'>
                    <h4>🏙️ Top Performing City</h4>
                    <p><b>{top_performing_city}</b></p>
                    <p>Revenue: {format_currency(top_city_sales)}</p>
                </div>
                """, unsafe_allow_html=True)

            with insights_col2:
                best_category = df.groupby("Product_Category")["Final_Sales_Amount"].sum().idxmax()
                best_category_sales = df.groupby("Product_Category")["Final_Sales_Amount"].sum().max()

                st.markdown(f"""
                <div class='insight-card'>
                    <h4>🏆 Best Category</h4>
                    <p><b>{best_category}</b></p>
                    <p>Revenue: {format_currency(best_category_sales)}</p>
                </div>
                """, unsafe_allow_html=True)

            with insights_col3:
                avg_discount_rate = df["Discount_Percent"].mean()

                st.markdown(f"""
                <div class='insight-card'>
                    <h4>💰 Discount Strategy</h4>
                    <p><b>{format_percentage(avg_discount_rate)}</b></p>
                    <p>Average Discount Rate</p>
                </div>
                """, unsafe_allow_html=True)

            except Exception as tab6_error:
                st.markdown(f"<div class='alert-box alert-error'>⚠️ Advanced Analytics temporarily unavailable. Error: {str(tab6_error)}</div>", unsafe_allow_html=True)
                st.info("Other tabs remain fully functional. Please check your data or contact support.")

        st.markdown("---")
        st.markdown("<p style='text-align:center; color:#6b7280; font-size:14px;'>Dashboard built with Streamlit | Data analytics powered by Python</p>", unsafe_allow_html=True)

    except Exception as e:
        st.markdown(f"<div class='alert-box alert-error'>❌ <b>Error:</b> {str(e)}</div>", unsafe_allow_html=True)
        st.error("Please check your data format and try again. Required columns: Order_Date, Customer_Name, City, Product_Category, Product_Name, Quantity, Unit_Price, Discount_Percent, Order_Status, Delivery_Days")

else:
    st.markdown("<div class='alert-box alert-info'>ℹ️ Please upload an Excel file to begin analysis</div>", unsafe_allow_html=True)

    st.markdown("""
    ### Required Data Format

    Your Excel file should contain the following columns:
    - **Order_Date**: Date of the order
    - **Customer_Name**: Name of the customer
    - **City**: City of the customer
    - **Product_Category**: Category of the product
    - **Product_Name**: Name of the product
    - **Quantity**: Number of items ordered
    - **Unit_Price**: Price per unit
    - **Discount_Percent**: Discount percentage applied
    - **Order_Status**: Status of the order (Delivered, Cancelled, Returned, etc.)
    - **Delivery_Days**: Number of days for delivery
    """)
