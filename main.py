import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Enterprise Sales Intelligence Platform",
    page_icon="📊",
    layout="wide"
)

# -------------------------------------------------
# CSS
# -------------------------------------------------
st.markdown("""
<style>
body {background:#f2f5f9;}

.metric-card{
background:#ffffff;
padding:18px;
border-radius:14px;
text-align:center;
box-shadow:0px 4px 12px rgba(0,0,0,0.08);
}

.metric-value{
font-size:34px;
font-weight:700;
color:#1f2937;
}

.metric-label{
font-size:14px;
color:#6b7280;
letter-spacing:0.5px;
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
st.markdown("<h1 style='text-align:center'>🚀 Enterprise Sales Intelligence Platform</h1>", unsafe_allow_html=True)

# -------------------------------------------------
# UPLOAD
# -------------------------------------------------
file = st.file_uploader("Upload Excel File", type="xlsx")

if file:

    df = pd.read_excel(file)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"])

    # -------------------------------------------------
    # FEATURE ENGINEERING
    # -------------------------------------------------
    df["Clean_Customer_Name"] = df["Customer_Name"].str.strip().str.title()
    df["Customer_Name_Upper"] = df["Customer_Name"].str.upper()
    df["Product_Category_Lower"] = df["Product_Category"].str.lower()

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

    df["Order_Summary"] = df["Customer_Name"]+" - "+df["Product_Name"]+" - "+df["City"]

    df["Month"] = df["Order_Date"].dt.to_period("M").astype(str)

    # -------------------------------------------------
    # SIDEBAR FILTERS
    # -------------------------------------------------
    st.sidebar.header("🔎 Filters")

    city = st.sidebar.multiselect("City", df["City"].unique())
    category = st.sidebar.multiselect("Category", df["Product_Category"].unique())
    date_range = st.sidebar.date_input("Date Range",
        [df["Order_Date"].min(),df["Order_Date"].max()])

    if city:
        df = df[df["City"].isin(city)]
    if category:
        df = df[df["Product_Category"].isin(category)]

    df = df[(df["Order_Date"]>=pd.to_datetime(date_range[0])) &
            (df["Order_Date"]<=pd.to_datetime(date_range[1]))]

    # -------------------------------------------------
    # KPIs
    # -------------------------------------------------
    total_sales = df["Final_Sales_Amount"].sum()
    avg_sales = df["Final_Sales_Amount"].mean()
    total_orders = len(df)
    fast_pct = (df["Fast_Delivery_Flag"]=="Fast Delivery").mean()*100

    k1,k2,k3,k4 = st.columns(4)
    k1.markdown(f"<div class='metric-card'><div class='metric-value'>{total_sales:,.0f}</div><div class='metric-label'>Total Sales</div></div>",unsafe_allow_html=True)
    k2.markdown(f"<div class='metric-card'><div class='metric-value'>{avg_sales:,.0f}</div><div class='metric-label'>Avg Order Value</div></div>",unsafe_allow_html=True)
    k3.markdown(f"<div class='metric-card'><div class='metric-value'>{total_orders}</div><div class='metric-label'>Total Orders</div></div>",unsafe_allow_html=True)
    k4.markdown(f"<div class='metric-card'><div class='metric-value'>{fast_pct:.1f}%</div><div class='metric-label'>Fast Delivery %</div></div>",unsafe_allow_html=True)

    # -------------------------------------------------
    # TABS
    # -------------------------------------------------
    tab1,tab2,tab3,tab4 = st.tabs(["📈 Overview","🏆 Performance","🚚 Delivery","📋 Data"])

    # ------------------- OVERVIEW -------------------
    with tab1:
        col1,col2 = st.columns(2)

        fig1 = px.line(df.groupby("Month")["Final_Sales_Amount"].sum().reset_index(),
                       x="Month",y="Final_Sales_Amount",title="Monthly Sales Trend")
        col1.plotly_chart(fig1,use_container_width=True)

        fig2 = px.bar(df,x="Order_Value_Category",title="Order Value Segments",
                      color="Order_Value_Category")
        col2.plotly_chart(fig2,use_container_width=True)

    # ------------------- PERFORMANCE -------------------
    with tab2:
        top_n = st.slider("Top N Customers",5,20,10)

        top_customers = df.groupby("Clean_Customer_Name")["Final_Sales_Amount"].sum()\
                           .sort_values(ascending=False).head(top_n).reset_index()

        fig3 = px.bar(top_customers,x="Final_Sales_Amount",
                      y="Clean_Customer_Name",orientation="h",
                      title="Top Customers")
        st.plotly_chart(fig3,use_container_width=True)

        heat = pd.pivot_table(df,values="Final_Sales_Amount",
                               index="City",columns="Product_Category",
                               aggfunc="sum",fill_value=0)

        fig4 = px.imshow(heat,text_auto=True,title="City × Category Sales Heatmap")
        st.plotly_chart(fig4,use_container_width=True)

    # ------------------- DELIVERY -------------------
    with tab3:
        fig5 = px.pie(df,names="Fast_Delivery_Flag",title="Delivery Speed")
        st.plotly_chart(fig5)

        fig6 = px.bar(df,x="Delivery_Status",title="Delivery Status")
        st.plotly_chart(fig6)

    # ------------------- DATA -------------------
    with tab4:
        st.dataframe(df)
        df.to_excel("processed_data.xlsx",index=False)
        with open("processed_data.xlsx","rb") as f:
            st.download_button("⬇ Download Processed File",f,"processed_data.xlsx")

else:
    st.info("Upload an Excel file to start")
