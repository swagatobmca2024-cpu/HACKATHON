import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# ----------------------------------
# PAGE CONFIG
# ----------------------------------
st.set_page_config(page_title="Sales Data Processing App", layout="wide")

st.title("📊 Sales Data Cleaning & Analysis App")

# ----------------------------------
# FILE UPLOAD
# ----------------------------------
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    df = pd.read_excel(uploaded_file)

    st.subheader("🔹 Raw Data Preview")
    st.dataframe(df.head())

    # ----------------------------------
    # 1. Clean_Customer_Name
    # ----------------------------------
    df["Clean_Customer_Name"] = (
        df["Customer_Name"]
        .astype(str)
        .str.strip()
        .str.title()
    )

    # ----------------------------------
    # 2. Customer_Name_Upper
    # ----------------------------------
    df["Customer_Name_Upper"] = (
        df["Customer_Name"]
        .astype(str)
        .str.upper()
    )

    # ----------------------------------
    # 3. Product_Category_Lower
    # ----------------------------------
    df["Product_Category_Lower"] = (
        df["Product_Category"]
        .astype(str)
        .str.lower()
    )

    # ----------------------------------
    # 4. Total_Amount (Before Discount)
    # ----------------------------------
    df["Total_Amount"] = df["Quantity"] * df["Unit_Price"]

    # ----------------------------------
    # 5. Discount_Amount
    # ----------------------------------
    df["Discount_Amount"] = (
        df["Total_Amount"] * df["Discount_Percent"] / 100
    )

    # ----------------------------------
    # 6. Final_Sales_Amount
    # ----------------------------------
    df["Final_Sales_Amount"] = (
        df["Total_Amount"] - df["Discount_Amount"]
    )

    # ----------------------------------
    # 7. Order_Value_Category
    # ----------------------------------
    def order_value_category(x):
        if x >= 30000:
            return "High Value"
        elif x >= 10000:
            return "Medium Value"
        else:
            return "Low Value"

    df["Order_Value_Category"] = df["Final_Sales_Amount"].apply(order_value_category)

    # ----------------------------------
    # 8. Delivery_Status
    # ----------------------------------
    def delivery_status(status):
        if status == "Cancelled":
            return "Not Delivered"
        elif status == "Returned":
            return "Returned"
        else:
            return "Delivered"

    df["Delivery_Status"] = df["Order_Status"].apply(delivery_status)

    # ----------------------------------
    # 9. Fast_Delivery_Flag
    # ----------------------------------
    def fast_delivery(days):
        if pd.isna(days):
            return "NA"
        elif days <= 3:
            return "Fast Delivery"
        else:
            return "Normal Delivery"

    df["Fast_Delivery_Flag"] = df["Delivery_Days"].apply(fast_delivery)

    # ----------------------------------
    # 10. Average Final Sales Amount
    # ----------------------------------
    avg_sales = df["Final_Sales_Amount"].mean()

    # ----------------------------------
    # 11. Min & Max Final Sales Amount
    # ----------------------------------
    min_sales = df["Final_Sales_Amount"].min()
    max_sales = df["Final_Sales_Amount"].max()

    # ----------------------------------
    # 12. Total Transactions
    # ----------------------------------
    total_transactions = len(df)

    # ----------------------------------
    # 13. Blank Delivery Days Count
    # ----------------------------------
    blank_delivery_days = df["Delivery_Days"].isna().sum()

    # ----------------------------------
    # 14. Order_Summary
    # ----------------------------------
    df["Order_Summary"] = (
        df["Customer_Name"]
        + " – "
        + df["Product_Name"]
        + " – "
        + df["City"]
    )

    # ----------------------------------
    # DISPLAY CLEANED DATA
    # ----------------------------------
    st.subheader("✅ Cleaned & Enhanced Dataset")
    st.dataframe(df)

    # ----------------------------------
    # KPI METRICS
    # ----------------------------------
    st.subheader("📌 Key Metrics")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Avg Final Sales", f"{avg_sales:,.2f}")
    col2.metric("Min Final Sales", f"{min_sales:,.2f}")
    col3.metric("Max Final Sales", f"{max_sales:,.2f}")
    col4.metric("Total Transactions", total_transactions)

    st.metric("Transactions with Blank Delivery Days", blank_delivery_days)

    # ----------------------------------
    # PLOTLY CHARTS
    # ----------------------------------

    st.subheader("📈 Visualizations")

    # Order Value Category Count
    fig1 = px.bar(
        df,
        x="Order_Value_Category",
        title="Orders by Value Category",
        color="Order_Value_Category"
    )
    st.plotly_chart(fig1)

    # Final Sales Distribution
    fig2 = px.histogram(
        df,
        x="Final_Sales_Amount",
        nbins=20,
        title="Final Sales Amount Distribution"
    )
    st.plotly_chart(fig2)

    # Fast Delivery Flag
    fig3 = px.pie(
        df,
        names="Fast_Delivery_Flag",
        title="Fast vs Normal Delivery"
    )
    st.plotly_chart(fig3)

    # ----------------------------------
    # DOWNLOAD RESULT
    # ----------------------------------
    st.subheader("⬇ Download Processed File")

    output_file = "processed_sales_data.xlsx"
    df.to_excel(output_file, index=False)

    with open(output_file, "rb") as file:
        st.download_button(
            label="Download Excel File",
            data=file,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Please upload an Excel file to begin.")
