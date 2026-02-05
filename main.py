import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import matplotlib.pyplot as plt

# -------------------------------------------------
# PAGE SETTINGS
# -------------------------------------------------
st.set_page_config(
    page_title="Advanced Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# CUSTOM CSS
# -------------------------------------------------
st.markdown("""
<style>

body {
background-color: #f4f6f9;
}

.metric-card {
background: white;
padding: 20px;
border-radius: 12px;
box-shadow: 0px 2px 8px rgba(0,0,0,0.1);
text-align:center;
}

.metric-value {
font-size:32px;
font-weight:bold;
color:#2c3e50;
}

.metric-label {
font-size:16px;
color:gray;
}

.section-title {
font-size:24px;
font-weight:bold;
margin-top:20px;
margin-bottom:10px;
color:#34495e;
}

</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# TITLE
# -------------------------------------------------
st.markdown("<h1 style='text-align:center'>🚀 Advanced Sales Analytics Dashboard</h1>", unsafe_allow_html=True)

# -------------------------------------------------
# FILE UPLOAD
# -------------------------------------------------
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file:

    df = pd.read_excel(file)

    # -------------------------------------------------
    # FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = df["Customer_Name"].astype(str).str.strip().str.title()
    df["Customer_Name_Upper"] = df["Customer_Name"].astype(str).str.upper()
    df["Product_Category_Lower"] = df["Product_Category"].astype(str).str.lower()

    df["Total_Amount"] = df["Quantity"] * df["Unit_Price"]
    df["Discount_Amount"] = df["Total_Amount"] * df["Discount_Percent"] / 100
    df["Final_Sales_Amount"] = df["Total_Amount"] - df["Discount_Amount"]

    def value_cat(x):
        if x >= 30000:
            return "High Value"
        elif x >= 10000:
            return "Medium Value"
        else:
            return "Low Value"

    df["Order_Value_Category"] = df["Final_Sales_Amount"].apply(value_cat)

    def delivery_status(x):
        if x == "Cancelled":
            return "Not Delivered"
        elif x == "Returned":
            return "Returned"
        else:
            return "Delivered"

    df["Delivery_Status"] = df["Order_Status"].apply(delivery_status)

    def fast_flag(x):
        if pd.isna(x):
            return "NA"
        elif x <= 3:
            return "Fast Delivery"
        else:
            return "Normal Delivery"

    df["Fast_Delivery_Flag"] = df["Delivery_Days"].apply(fast_flag)

    df["Order_Summary"] = df["Customer_Name"] + " - " + df["Product_Name"] + " - " + df["City"]

    # -------------------------------------------------
    # SIDEBAR FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔎 Filters")

    city_filter = st.sidebar.multiselect("Select City", df["City"].unique())
    category_filter = st.sidebar.multiselect("Select Product Category", df["Product_Category"].unique())

    if city_filter:
        df = df[df["City"].isin(city_filter)]

    if category_filter:
        df = df[df["Product_Category"].isin(category_filter)]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    avg_sales = round(df["Final_Sales_Amount"].mean(),2)
    min_sales = round(df["Final_Sales_Amount"].min(),2)
    max_sales = round(df["Final_Sales_Amount"].max(),2)
    total_txn = len(df)
    blank_delivery = df["Delivery_Days"].isna().sum()

    st.markdown("<div class='section-title'>📌 Key Metrics</div>", unsafe_allow_html=True)

    c1,c2,c3,c4,c5 = st.columns(5)

    c1.markdown(f"<div class='metric-card'><div class='metric-value'>{avg_sales}</div><div class='metric-label'>Avg Sales</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='metric-card'><div class='metric-value'>{min_sales}</div><div class='metric-label'>Min Sales</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='metric-card'><div class='metric-value'>{max_sales}</div><div class='metric-label'>Max Sales</div></div>", unsafe_allow_html=True)
    c4.markdown(f"<div class='metric-card'><div class='metric-value'>{total_txn}</div><div class='metric-label'>Transactions</div></div>", unsafe_allow_html=True)
    c5.markdown(f"<div class='metric-card'><div class='metric-value'>{blank_delivery}</div><div class='metric-label'>Blank Delivery Days</div></div>", unsafe_allow_html=True)

    # -------------------------------------------------
    # CHARTS
    # -------------------------------------------------
    st.markdown("<div class='section-title'>📊 Visual Insights</div>", unsafe_allow_html=True)

    col1,col2 = st.columns(2)

    fig1 = px.bar(df, x="Order_Value_Category", title="Orders by Value Category", color="Order_Value_Category")
    col1.plotly_chart(fig1, use_container_width=True)

    fig2 = px.histogram(df, x="Final_Sales_Amount", title="Final Sales Distribution")
    col2.plotly_chart(fig2, use_container_width=True)

    col3,col4 = st.columns(2)

    fig3 = px.pie(df, names="Fast_Delivery_Flag", title="Delivery Speed")
    col3.plotly_chart(fig3, use_container_width=True)

    city_sales = df.groupby("City")["Final_Sales_Amount"].sum().reset_index()
    fig4 = px.bar(city_sales, x="City", y="Final_Sales_Amount", title="Sales by City")
    col4.plotly_chart(fig4, use_container_width=True)

    # -------------------------------------------------
    # DATA TABLE
    # -------------------------------------------------
    st.markdown("<div class='section-title'>📋 Processed Data</div>", unsafe_allow_html=True)
    st.dataframe(df)

    # -------------------------------------------------
    # DOWNLOAD
    # -------------------------------------------------
    df.to_excel("processed_data.xlsx", index=False)
    with open("processed_data.xlsx","rb") as f:
        st.download_button("⬇ Download Processed File", f, "processed_data.xlsx")

else:
    st.info("Please upload Excel file to start")
