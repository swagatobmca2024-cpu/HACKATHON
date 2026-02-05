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
font-size:30px;
font-weight:700;
color:#1f2937;
}

.metric-label{
font-size:13px;
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

    # -------------------------------------------------
    # LOAD DATA
    # -------------------------------------------------
    df = pd.read_excel(file)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], dayfirst=True, errors="coerce")

    # -------------------------------------------------
    # FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = (
        df["Customer_Name"].astype(str)
        .str.strip()
        .str.replace(r"\s+"," ", regex=True)
        .str.title()
    )

    df["Customer_Name_Upper"] = df["Customer_Name"].astype(str).str.upper()
    df["Product_Category_Lower"] = df["Product_Category"].astype(str).str.lower()

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
        df["Customer_Name"].astype(str) + " – " +
        df["Product_Name"].astype(str) + " – " +
        df["City"].astype(str)
    )

    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    # -------------------------------------------------
    # FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔎 Filters")
    city = st.sidebar.multiselect("City", df["City"].unique())
    category = st.sidebar.multiselect("Category", df["Product_Category"].unique())

    if city:
        df = df[df["City"].isin(city)]
    if category:
        df = df[df["Product_Category"].isin(category)]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    total_sales = df["Final_Sales_Amount"].sum()
    total_orders = len(df)
    total_items = df["Quantity"].sum()
    avg_order = df["Final_Sales_Amount"].mean()
    discount_given = df["Discount_Amount"].sum()
    customers = df["Clean_Customer_Name"].nunique()

    min_sale = df["Final_Sales_Amount"].min()
    max_sale = df["Final_Sales_Amount"].max()
    blank_delivery = df["Delivery_Days"].isna().sum()

    st.markdown("<div class='section-title'>📌 Key Metrics</div>", unsafe_allow_html=True)

    k1,k2,k3,k4,k5,k6,k7,k8,k9 = st.columns(9)

    k1.markdown(f"<div class='metric-card'><div class='metric-value'>₹{total_sales:,.0f}</div><div class='metric-label'>Total Sales</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><div class='metric-value'>{total_orders}</div><div class='metric-label'>Orders</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><div class='metric-value'>{total_items}</div><div class='metric-label'>Items Sold</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><div class='metric-value'>₹{avg_order:,.0f}</div><div class='metric-label'>Avg Order</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='metric-card'><div class='metric-value'>₹{discount_given:,.0f}</div><div class='metric-label'>Discount</div></div>", unsafe_allow_html=True)
    k6.markdown(f"<div class='metric-card'><div class='metric-value'>{customers}</div><div class='metric-label'>Customers</div></div>", unsafe_allow_html=True)
    k7.markdown(f"<div class='metric-card'><div class='metric-value'>₹{min_sale:,.0f}</div><div class='metric-label'>Min Sale</div></div>", unsafe_allow_html=True)
    k8.markdown(f"<div class='metric-card'><div class='metric-value'>₹{max_sale:,.0f}</div><div class='metric-label'>Max Sale</div></div>", unsafe_allow_html=True)
    k9.markdown(f"<div class='metric-card'><div class='metric-value'>{blank_delivery}</div><div class='metric-label'>Blank Delivery Days</div></div>", unsafe_allow_html=True)

    # -------------------------------------------------
    # TABS
    # -------------------------------------------------
    tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs(
        ["📊 Overview","📈 Trends","🏆 Products","👥 Customers","🚚 Delivery","💳 Payments & Channels"]
    )

    # ---------------- OVERVIEW ----------------
    with tab1:
        col1,col2 = st.columns(2)

        ovc = df["Order_Value_Category"].value_counts().reset_index()
        ovc.columns=["Category","Count"]
        col1.plotly_chart(px.bar(ovc,x="Category",y="Count",title="Orders by Value Category"), use_container_width=True)

        city_sales = df.groupby("City")["Final_Sales_Amount"].sum().reset_index()
        col2.plotly_chart(px.bar(city_sales,x="City",y="Final_Sales_Amount",title="Sales by City"), use_container_width=True)

        heat = df.pivot_table(index="City", columns="Product_Category", values="Final_Sales_Amount", aggfunc="sum", fill_value=0)
        st.plotly_chart(px.imshow(heat, text_auto=True, aspect="auto", title="City vs Category Heatmap"), use_container_width=True)

    # ---------------- TRENDS ----------------
    with tab2:
        monthly = df.groupby("Month")["Final_Sales_Amount"].sum().reset_index()
        st.plotly_chart(px.line(monthly,x="Month",y="Final_Sales_Amount",markers=True,title="Monthly Sales Trend"), use_container_width=True)

    # ---------------- PRODUCTS ----------------
    with tab3:
        prod = df.groupby("Product_Name")["Final_Sales_Amount"].sum().reset_index().sort_values(by="Final_Sales_Amount",ascending=False).head(10)
        st.plotly_chart(px.bar(prod,x="Final_Sales_Amount",y="Product_Name",orientation="h",title="Top Products"), use_container_width=True)

    # ---------------- CUSTOMERS ----------------
    with tab4:
        cust = df.groupby("Clean_Customer_Name")["Final_Sales_Amount"].sum().reset_index().sort_values(by="Final_Sales_Amount",ascending=False).head(10)
        st.plotly_chart(px.bar(cust,x="Final_Sales_Amount",y="Clean_Customer_Name",orientation="h",title="Top Customers"), use_container_width=True)

    # ---------------- DELIVERY ----------------
    with tab5:
        st.plotly_chart(px.pie(df["Fast_Delivery_Flag"].value_counts().reset_index(),
                               names="index", values="Fast_Delivery_Flag", title="Delivery Speed"), use_container_width=True)

        st.plotly_chart(px.bar(df["Delivery_Status"].value_counts().reset_index(),
                               x="index", y="Delivery_Status", title="Delivery Status"), use_container_width=True)

    # ---------------- PAYMENTS & CHANNELS ----------------
    with tab6:
        col1,col2 = st.columns(2)

        pay = df["Payment_Mode"].value_counts().reset_index()
        col1.plotly_chart(px.bar(pay,x="index",y="Payment_Mode",title="Orders by Payment Mode"), use_container_width=True)

        channel = df.groupby("Sales_Channel")["Final_Sales_Amount"].sum().reset_index()
        col2.plotly_chart(px.pie(channel,names="Sales_Channel",values="Final_Sales_Amount",title="Sales by Channel"), use_container_width=True)

else:
    st.info("Upload an Excel file to begin")
