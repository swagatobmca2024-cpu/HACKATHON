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
box-shadow:0 4px 10px rgba(0,0,0,0.15);
}

.metric-value{
font-size:32px;
font-weight:700;
color:#111827;
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

    # -------------------------------------------------
    # LOAD DATA
    # -------------------------------------------------
    df = pd.read_excel(file)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], dayfirst=True, errors="coerce")

    # -------------------------------------------------
    # FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = (
        df["Customer_Name"]
        .astype(str)
        .str.strip()
        .str.replace(r"\s+"," ",regex=True)
        .str.title()
    )

    df["Total_Amount"] = df["Quantity"] * df["Unit_Price"]
    df["Discount_Amount"] = df["Total_Amount"] * df["Discount_Percent"] / 100
    df["Final_Sales_Amount"] = df["Total_Amount"] - df["Discount_Amount"]

    df["Order_Value_Category"] = np.where(
        df["Final_Sales_Amount"] >= 30000,"High Value",
        np.where(df["Final_Sales_Amount"] >= 10000,"Medium Value","Low Value")
    )

    df["Delivery_Status"] = np.where(
        df["Order_Status"]=="Cancelled","Not Delivered",
        np.where(df["Order_Status"]=="Returned","Returned","Delivered")
    )

    df["Fast_Delivery_Flag"] = np.where(
        df["Delivery_Days"].isna(),"NA",
        np.where(df["Delivery_Days"]<=3,"Fast Delivery","Normal Delivery")
    )

    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    # -------------------------------------------------
    # SIDEBAR FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔎 Filters")

    city = st.sidebar.multiselect("City", df["City"].unique())
    category = st.sidebar.multiselect("Category", df["Product_Category"].unique())
    payment = st.sidebar.multiselect("Payment Mode", df["Payment_Mode"].unique())

    if city:
        df = df[df["City"].isin(city)]
    if category:
        df = df[df["Product_Category"].isin(category)]
    if payment:
        df = df[df["Payment_Mode"].isin(payment)]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    total_sales = df["Final_Sales_Amount"].sum()
    total_orders = len(df)
    total_items = df["Quantity"].sum()
    avg_order = total_sales / total_orders if total_orders else 0
    customers = df["Clean_Customer_Name"].nunique()
    fast_pct = (df["Fast_Delivery_Flag"]=="Fast Delivery").mean()*100

    st.markdown("<div class='section-title'>📌 Key Metrics</div>", unsafe_allow_html=True)

    k1,k2,k3,k4,k5,k6 = st.columns(6)

    k1.markdown(f"<div class='metric-card'><div class='metric-value'>₹{total_sales:,.0f}</div><div class='metric-label'>Total Sales</div></div>", unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><div class='metric-value'>{total_orders}</div><div class='metric-label'>Orders</div></div>", unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><div class='metric-value'>{total_items}</div><div class='metric-label'>Items Sold</div></div>", unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><div class='metric-value'>₹{avg_order:,.0f}</div><div class='metric-label'>Avg Order Value</div></div>", unsafe_allow_html=True)
    k5.markdown(f"<div class='metric-card'><div class='metric-value'>{customers}</div><div class='metric-label'>Customers</div></div>", unsafe_allow_html=True)
    k6.markdown(f"<div class='metric-card'><div class='metric-value'>{fast_pct:.1f}%</div><div class='metric-label'>Fast Delivery %</div></div>", unsafe_allow_html=True)

    # -------------------------------------------------
    # TABS
    # -------------------------------------------------
    tab1,tab2,tab3,tab4,tab5,tab6 = st.tabs(
        ["📊 Overview","📈 Trends","🏆 Products","👥 Customers",
         "💳 Payments","🚚 Delivery"]
    )

    # ---------------- OVERVIEW ----------------
    with tab1:

        col1,col2 = st.columns(2)

        ovc = df["Order_Value_Category"].value_counts().reset_index()
        ovc.columns=["Category","Orders"]

        fig1 = px.bar(
            ovc,x="Category",y="Orders",
            text="Orders",
            title="Orders by Value Category"
        )
        col1.plotly_chart(fig1,use_container_width=True)

        city_orders = df.groupby("City").size().reset_index(name="Orders")

        fig2 = px.bar(
            city_orders,x="City",y="Orders",
            title="Orders by City",
            text="Orders"
        )
        col2.plotly_chart(fig2,use_container_width=True)

        heat = df.pivot_table(
            index="City",
            columns="Product_Category",
            values="Final_Sales_Amount",
            aggfunc="sum",
            fill_value=0
        )

        fig_heat = px.imshow(
            heat,
            text_auto=True,
            color_continuous_scale="Turbo",
            aspect="auto",
            title="City vs Category Sales Heatmap"
        )
        fig_heat.update_layout(height=650)

        st.plotly_chart(fig_heat,use_container_width=True)

    # ---------------- TRENDS ----------------
    with tab2:
        monthly = df.groupby("Month")["Final_Sales_Amount"].sum().reset_index()

        fig3 = px.line(
            monthly,x="Month",y="Final_Sales_Amount",
            markers=True,title="Monthly Revenue Trend"
        )
        st.plotly_chart(fig3,use_container_width=True)

    # ---------------- PRODUCTS ----------------
    with tab3:
        prod = (
            df.groupby("Product_Name")["Final_Sales_Amount"]
            .sum().reset_index()
            .sort_values(by="Final_Sales_Amount",ascending=False)
            .head(10)
        )

        fig4 = px.bar(
            prod,x="Final_Sales_Amount",y="Product_Name",
            orientation="h",
            title="Top Products by Sales"
        )
        st.plotly_chart(fig4,use_container_width=True)

    # ---------------- CUSTOMERS ----------------
    with tab4:
        cust = (
            df.groupby("Clean_Customer_Name")
            .agg(
                Orders=("Product_Name","count"),
                Sales=("Final_Sales_Amount","sum")
            )
            .reset_index()
            .sort_values(by="Sales",ascending=False)
            .head(15)
        )

        fig5 = px.bar(
            cust,x="Sales",y="Clean_Customer_Name",
            orientation="h",
            text="Orders",
            title="Top Customers (Sales + Orders)"
        )
        st.plotly_chart(fig5,use_container_width=True)

    # ---------------- PAYMENTS ----------------
    with tab5:
        pay = df["Payment_Mode"].value_counts().reset_index()
        pay.columns=["Payment_Mode","Orders"]

        fig6 = px.pie(
            pay,names="Payment_Mode",
            values="Orders",
            title="Orders by Payment Mode"
        )
        st.plotly_chart(fig6)

        channel = df["Sales_Channel"].value_counts().reset_index()
        channel.columns=["Sales_Channel","Orders"]

        fig7 = px.bar(
            channel,x="Sales_Channel",y="Orders",
            text="Orders",
            title="Orders by Sales Channel"
        )
        st.plotly_chart(fig7,use_container_width=True)

    # ---------------- DELIVERY ----------------
    with tab6:
        fast = df["Fast_Delivery_Flag"].value_counts().reset_index()
        fast.columns=["Type","Orders"]

        fig8 = px.pie(
            fast,names="Type",
            values="Orders",
            title="Delivery Speed"
        )
        st.plotly_chart(fig8)

        delv = df["Delivery_Status"].value_counts().reset_index()
        delv.columns=["Status","Orders"]

        fig9 = px.bar(
            delv,x="Status",y="Orders",
            text="Orders",
            title="Delivery Status"
        )
        st.plotly_chart(fig9,use_container_width=True)

else:
    st.info("Upload an Excel file to begin")
