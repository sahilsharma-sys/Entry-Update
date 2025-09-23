import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import os
from datetime import datetime, timedelta

# === File Path ===
file_path = "crm_allocation.xlsx"

# === Load existing or create new ===
if os.path.exists(file_path):
    df = pd.read_excel(file_path)

    # ✅ Force Date column to datetime (fix for old entries)
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
else:
    df = pd.DataFrame(columns=["Date", "Merchant", "Courier", "Remarks"])
    df.to_excel(file_path, index=False, engine="openpyxl")

# === Streamlit UI ===
st.set_page_config(page_title="CRM Allocation Dashboard", layout="wide")
st.title("📦 CRM Allocation Dashboard")

# Sidebar Navigation
menu = st.sidebar.radio("📌 Menu", ["Add Entry", "Dashboard", "Logs"])

# === Add Entry Page ===
if menu == "Add Entry":
    st.header("➕ Add New Entry")

    merchant = st.text_input("Merchant Name")
    couriers = [
        "ATS", "NimbusPost", "Blue Dart Direct", "Shipyaari", "Delhivery",
        "Ecom Express", "Ekart", "Shadowfax", "Xpressbees", "DTDC"
    ]
    courier = st.selectbox("Courier", couriers)
    remarks = st.text_area("Remarks")

    if st.button("Save Entry"):
        if merchant and courier:
            new_data = pd.DataFrame([{
                "Date": datetime.now(),
                "Merchant": merchant,
                "Courier": courier,
                "Remarks": remarks
            }])
            df = pd.concat([df, new_data], ignore_index=True)
            df.to_excel(file_path, index=False, engine="openpyxl")
            st.success("✅ Entry Saved Successfully!")
        else:
            st.error("⚠️ Please fill all required fields.")

# === Dashboard Page ===
elif menu == "Dashboard":
    st.header("📊 Dashboard Overview")

    if not df.empty:
        # ✅ Last Update
        if df["Date"].notnull().any():
            last_update = df["Date"].max().strftime("%Y-%m-%d %H:%M:%S")
            st.info(f"🕒 Last Update: {last_update}")

        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total Merchants", df["Merchant"].nunique())

        with col2:
            st.metric("Total Couriers Used", df["Courier"].nunique())

        with col3:
            st.metric("Total Entries", len(df))

        # === Highlights ===
        st.subheader("🌟 Highlights")
        courier_counts = df["Courier"].value_counts().head(5)
        st.write("**Top Couriers:**")
        st.bar_chart(courier_counts)

        merchant_counts = df["Merchant"].value_counts().head(5)
        st.write("**Top Merchants:**")
        st.bar_chart(merchant_counts)

        # === Weekly Summary ===
        st.subheader("📅 Weekly Summary")
        last_week = datetime.now() - timedelta(days=7)
        weekly_data = df[df["Date"].notnull() & (df["Date"] >= last_week)]

        if not weekly_data.empty:
            weekly_summary = weekly_data.groupby("Courier").size()
            st.bar_chart(weekly_summary)
        else:
            st.warning("No data available for this week.")
    else:
        st.warning("No data available yet. Please add some entries.")

# === Logs Page ===
elif menu == "Logs":
    st.header("📜 All Logs")
    if not df.empty:
        st.dataframe(df.sort_values(by="Date", ascending=False), use_container_width=True)

        # Export Options
        col1, col2 = st.columns(2)

        with col1:
            excel_data = df.to_excel(index=False, engine="openpyxl")
            st.download_button("⬇️ Download Logs (Excel)", excel_data, file_name="crm_logs.xlsx")

        with col2:
            csv_data = df.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download Logs (CSV)", csv_data, file_name="crm_logs.csv")
    else:
        st.warning("No logs to display.")
