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
# THEME
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
.metric-value{font-size:30px;font-weight:700;color:#1f2937;}
.metric-label{font-size:14px;color:#6b7280;}
.section-title{font-size:26px;font-weight:700;margin:20px 0;}
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------
# TITLE
# -------------------------------------------------
st.markdown("<h1 style='text-align:center'>🚀 Enterprise Sales Intelligence Dashboard</h1>", unsafe_allow_html=True)

# -------------------------------------------------
# UPLOAD
# -------------------------------------------------
file = st.file_uploader("Upload Excel File", type=["xlsx"])

if file:

    # -------------------------------------------------
    # LOAD DATA
    # -------------------------------------------------
    df = pd.read_excel(file)

    df["Order_Date"] = pd.to_datetime(
        df["Order_Date"],
        dayfirst=True,
        errors="coerce"
    )

    df["Month"] = df["Order_Date"].dt.strftime("%Y-%m")

    # -------------------------------------------------
    # FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = (
        df["Customer_Name"].astype(str)
        .str.strip()
        .str.replace(r"\s+"," ",regex=True)
        .str.title()
    )

    df["Total_Amount"] = df["Quantity"] * df["Unit_Price"]
    df["Discount_Amount"] = df["Total_Amount"] * df["Discount_Percent"] / 100
    df["Final_Sales_Amount"] = df["Total_Amount"] - df["Discount_Amount"]

    df["Order_Value_Category"] = np.where(
        df["Final_Sales_Amount"]>=30000,"High Value",
        np.where(df["Final_Sales_Amount"]>=10000,"Medium Value","Low Value")
    )

    df["Delivery_Status"] = np.where(
        df["Order_Status"]=="Cancelled","Not Delivered",
        np.where(df["Order_Status"]=="Returned","Returned","Delivered")
    )

    df["Fast_Delivery_Flag"] = np.where(
        df["Delivery_Days"].isna(),"NA",
        np.where(df["Delivery_Days"]<=3,"Fast Delivery","Normal Delivery")
    )

    # -------------------------------------------------
    # FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔍 Filters")

    city = st.sidebar.multiselect("City", sorted(df["City"].unique()))
    category = st.sidebar.multiselect("Category", sorted(df["Product_Category"].unique()))
    channel = st.sidebar.multiselect("Sales Channel", sorted(df["Sales_Channel"].unique()))

    if city:
        df = df[df["City"].isin(city)]
    if category:
        df = df[df["Product_Category"].isin(category)]
    if channel:
        df = df[df["Sales_Channel"].isin(channel)]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    k1,k2,k3,k4,k5,k6 = st.columns(6)

    k1.markdown(f"<div class='metric-card'><div class='metric-value'>₹{df['Final_Sales_Amount'].sum():,.0f}</div><div class='metric-label'>Total Sales</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><div class='metric-value'>{len(df)}</div><div class='metric-label'>Orders</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><div class='metric-value'>{df['Clean_Customer_Name'].nunique()}</div><div class='metric-label'>Customers</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><div class='metric-value'>₹{df['Final_Sales_Amount'].mean():,.0f}</div><div class='metric-label'>Avg Order</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='metric-card'><div class='metric-value'>₹{df['Discount_Amount'].sum():,.0f}</div><div class='metric-label'>Discount Given</div></div>", unsafe_allow_html=True)
    k6.markdown(f"<div class='metric-card'><div class='metric-value'>{(df['Fast_Delivery_Flag']=='Fast Delivery').mean()*100:.1f}%</div><div class='metric-label'>Fast Delivery %</div></div>", unsafe_allow_html=True)

    # -------------------------------------------------
    # TABS
    # -------------------------------------------------
    tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs(
        ["📊 Overview","📈 Trends","🏆 Products","👥 Customers","🚚 Delivery","💳 Channels & Payments"]
    )

    # ---------------- OVERVIEW ----------------
    with tab1:

        col1,col2 = st.columns(2)

        fig1 = px.bar(
            df["Order_Value_Category"].value_counts().reset_index(),
            x="index",y="Order_Value_Category",
            title="Orders by Value Category"
        )
        col1.plotly_chart(fig1,use_container_width=True)

        fig2 = px.bar(
            df.groupby("City")["Final_Sales_Amount"].sum().reset_index(),
            x="City",y="Final_Sales_Amount",
            title="Sales by City"
        )
        col2.plotly_chart(fig2,use_container_width=True)

        heat = df.pivot_table(
            index="City",
            columns="Product_Category",
            values="Final_Sales_Amount",
            aggfunc="sum"
        )

        fig_heat = px.imshow(
            heat,
            color_continuous_scale="Turbo",
            aspect="auto",
            height=550,
            title="City vs Category Sales Heatmap"
        )

        st.plotly_chart(fig_heat,use_container_width=True)

    # ---------------- TRENDS ----------------
    with tab2:
        monthly = df.groupby("Month")["Final_Sales_Amount"].sum().reset_index()

        fig3 = px.line(
            monthly,
            x="Month",y="Final_Sales_Amount",
            markers=True,
            title="Monthly Revenue Trend"
        )

        st.plotly_chart(fig3,use_container_width=True)

    # ---------------- PRODUCTS ----------------
    with tab3:

        prod = df.groupby("Product_Name")["Final_Sales_Amount"].sum().reset_index()
        prod = prod.sort_values(by="Final_Sales_Amount",ascending=False).head(10)

        fig4 = px.bar(
            prod,
            x="Final_Sales_Amount",
            y="Product_Name",
            orientation="h",
            title="Top 10 Products by Revenue"
        )

        st.plotly_chart(fig4,use_container_width=True)

    # ---------------- CUSTOMERS ----------------
    with tab4:

        cust = df.groupby("Clean_Customer_Name").agg(
            Orders=("Order_ID","count"),
            Revenue=("Final_Sales_Amount","sum")
        ).reset_index()

        cust = cust.sort_values(by="Revenue",ascending=False).head(10)

        fig5 = px.bar(
            cust,
            x="Revenue",
            y="Clean_Customer_Name",
            orientation="h",
            title="Top Customers by Revenue"
        )

        st.plotly_chart(fig5,use_container_width=True)

    # ---------------- DELIVERY ----------------
    with tab5:

        fig6 = px.pie(
            df["Fast_Delivery_Flag"].value_counts().reset_index(),
            names="index",values="Fast_Delivery_Flag",
            title="Delivery Speed"
        )

        st.plotly_chart(fig6)

        fig7 = px.bar(
            df["Delivery_Status"].value_counts().reset_index(),
            x="index",y="Delivery_Status",
            title="Delivery Status"
        )

        st.plotly_chart(fig7)

    # ---------------- CHANNELS & PAYMENTS ----------------
    with tab6:

        col1,col2 = st.columns(2)

        fig8 = px.bar(
            df["Payment_Mode"].value_counts().reset_index(),
            x="index",y="Payment_Mode",
            title="Orders by Payment Mode"
        )
        col1.plotly_chart(fig8,use_container_width=True)

        fig9 = px.bar(
            df["Sales_Channel"].value_counts().reset_index(),
            x="index",y="Sales_Channel",
            title="Orders by Sales Channel"
        )
        col2.plotly_chart(fig9,use_container_width=True)

else:
    st.info("Upload an Excel file to begin")
