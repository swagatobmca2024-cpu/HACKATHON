import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Enterprise Sales Intelligence Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# CUSTOM CSS
# -------------------------------------------------
st.markdown("""
<style>
body {background:#f4f6f9;}

.metric-card{
background:white;
padding:18px;
border-radius:14px;
text-align:center;
box-shadow:0 4px 10px rgba(0,0,0,0.1);
}

.metric-value{
font-size:34px;
font-weight:700;
color:#1f2937;
}

.metric-label{
font-size:14px;
color:#6b7280;
}

.section-title{
font-size:26px;
font-weight:700;
margin:20px 0 10px 0;
color:#111827;
}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# TITLE
# -------------------------------------------------
st.markdown("<h1 style='text-align:center'>🚀 Enterprise Sales Intelligence Dashboard</h1>", unsafe_allow_html=True)

# -------------------------------------------------
# FILE UPLOAD
# -------------------------------------------------
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file:

    # 🔹 FIX DATE FORMAT
    df = pd.read_excel(file)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], dayfirst=True, errors="coerce")

    # -------------------------------------------------
    # CLEANING + FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = (
        df["Customer_Name"].astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.title()
    )

    df["Customer_Name_Upper"] = df["Customer_Name"].str.upper()
    df["Product_Category_Lower"] = df["Product_Category"].str.lower()

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

    df["Order_Summary"] = (
        df["Clean_Customer_Name"] + " - " +
        df["Product_Name"] + " - " +
        df["City"]
    )

    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    # -------------------------------------------------
    # SIDEBAR FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔎 Filters")

    city_filter = st.sidebar.multiselect("City", df["City"].unique())
    category_filter = st.sidebar.multiselect("Category", df["Product_Category"].unique())

    date_range = st.sidebar.date_input(
        "Order Date Range",
        [df["Order_Date"].min(), df["Order_Date"].max()]
    )

    if city_filter:
        df = df[df["City"].isin(city_filter)]

    if category_filter:
        df = df[df["Product_Category"].isin(category_filter)]

    df = df[(df["Order_Date"] >= pd.to_datetime(date_range[0])) &
            (df["Order_Date"] <= pd.to_datetime(date_range[1]))]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    total_sales = df["Final_Sales_Amount"].sum()
    avg_sales = df["Final_Sales_Amount"].mean()
    total_orders = len(df)
    fast_pct = (df["Fast_Delivery_Flag"] == "Fast Delivery").mean() * 100

    st.markdown("<div class='section-title'>📌 Key Metrics</div>", unsafe_allow_html=True)

    k1,k2,k3,k4 = st.columns(4)

    k1.markdown(f"<div class='metric-card'><div class='metric-value'>{total_sales:,.0f}</div><div class='metric-label'>Total Sales</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><div class='metric-value'>{avg_sales:,.0f}</div><div class='metric-label'>Avg Order Value</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><div class='metric-value'>{total_orders}</div><div class='metric-label'>Total Orders</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><div class='metric-value'>{fast_pct:.1f}%</div><div class='metric-label'>Fast Delivery %</div></div>", unsafe_allow_html=True)

    # -------------------------------------------------
    # TABS
    # -------------------------------------------------
    tab1, tab2, tab3, tab4, tab5 = st.tabs(
        ["📊 Overview","📈 Trends","🏆 Customers","🚚 Delivery","📋 Data"]
    )

    # ---------------- OVERVIEW ----------------
    with tab1:
        col1,col2 = st.columns(2)

        ovc = df["Order_Value_Category"].value_counts().reset_index()
        ovc.columns = ["Order_Value_Category","Count"]
        fig1 = px.bar(ovc, x="Order_Value_Category", y="Count",
                      text="Count", title="Orders by Value Category")
        col1.plotly_chart(fig1, use_container_width=True)

        city_sales = df.groupby("City")["Final_Sales_Amount"].sum().reset_index()
        fig2 = px.bar(city_sales, x="City", y="Final_Sales_Amount",
                      title="Sales by City")
        col2.plotly_chart(fig2, use_container_width=True)

    # ---------------- TRENDS ----------------
    with tab2:
        monthly = df.groupby("Month")["Final_Sales_Amount"].sum().reset_index()
        fig3 = px.line(monthly, x="Month", y="Final_Sales_Amount",
                       markers=True, title="Monthly Sales Trend")
        st.plotly_chart(fig3, use_container_width=True)

    # ---------------- CUSTOMERS ----------------
    with tab3:
        top_n = st.slider("Select Top N Customers", 5, 20, 10)
        top_customers = (
            df.groupby("Clean_Customer_Name")["Final_Sales_Amount"]
            .sum().sort_values(ascending=False)
            .head(top_n).reset_index()
        )

        fig4 = px.bar(top_customers, x="Final_Sales_Amount",
                      y="Clean_Customer_Name", orientation="h",
                      title="Top Customers by Sales")
        st.plotly_chart(fig4, use_container_width=True)

    # ---------------- DELIVERY ----------------
    with tab4:
        fast_counts = df["Fast_Delivery_Flag"].value_counts().reset_index()
        fast_counts.columns = ["Fast_Delivery_Flag","Count"]
        fig5 = px.pie(fast_counts, names="Fast_Delivery_Flag",
                      values="Count", title="Delivery Speed")
        st.plotly_chart(fig5)

        delivery_counts = df["Delivery_Status"].value_counts().reset_index()
        delivery_counts.columns = ["Delivery_Status","Count"]
        fig6 = px.bar(delivery_counts, x="Delivery_Status",
                      y="Count", text="Count",
                      title="Delivery Status")
        st.plotly_chart(fig6)

    # ---------------- DATA ----------------
    with tab5:
        st.dataframe(df)
        df.to_excel("processed_data.xlsx", index=False)
        with open("processed_data.xlsx","rb") as f:
            st.download_button("⬇ Download Processed Data", f, "processed_data.xlsx")

else:
    st.info("Upload an Excel file to begin")
