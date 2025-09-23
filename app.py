import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io

st.set_page_config(page_title="Allocation CRM", layout="centered")

# ---------------------------
# Config
# ---------------------------
FILE_PATH = "allocation_updates.xlsx"
COURIERS = [
    "ATS", "NimbusPost", "Blue Dart Direct", "Shipyaari", "Delhivery",
    "Ecom Express", "Ekart", "Shadowfax", "Xpressbees", "Blitz",
    "Professional", "GoSwift", "Pikndel", "DTDC"
]

# ---------------------------
# Load or Create File
# ---------------------------
if os.path.exists(FILE_PATH):
    df = pd.read_excel(FILE_PATH, engine="openpyxl")
else:
    df = pd.DataFrame(columns=["Date", "Merchant", "Courier", "Remarks"])
    df.to_excel(FILE_PATH, index=False, engine="openpyxl")

# Ensure Date is datetime.date
if not df.empty:
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

# Session-state for merchant history
if "merchant_history" not in st.session_state:
    st.session_state.merchant_history = df["Merchant"].dropna().unique().tolist()

# ---------------------------
# Tabs
# ---------------------------
tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "➕ Add Entry", "📄 Logs"])

# ---------------------------
# TAB 1: Dashboard
# ---------------------------
with tab1:
    st.header("📊 Allocation Dashboard ")

    if not df.empty:
        min_date, max_date = df["Date"].min(), df["Date"].max()
        date_range = st.date_input("📅 Filter by Date Range (Optional)", [min_date, max_date])
        filtered_df = df.copy()
        if len(date_range) == 2:
            filtered_df = filtered_df[(filtered_df["Date"] >= date_range[0]) &
                                      (filtered_df["Date"] <= date_range[1])]
    else:
        filtered_df = df.copy()

    # Summary Cards
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Updates", len(filtered_df))
    c2.metric("Unique Merchants", filtered_df["Merchant"].nunique())
    c3.metric("Unique Couriers Used", filtered_df["Courier"].nunique())

    if not filtered_df.empty:
        last_update = filtered_df["Date"].max().strftime("%d-%b-%y")
        st.info(f"🕒 Last Update: {last_update}")

        # Insights
        top_courier = filtered_df['Courier'].str.split(' \| ').explode().value_counts().idxmax()
        top_merchant = filtered_df['Merchant'].value_counts().idxmax()
        col1, col2 = st.columns(2)
        col1.metric("Most Used Courier", top_courier)
        col2.metric("Most Updated Merchant", top_merchant)

        # Top Merchants
        st.subheader("🛍️ Top 10 Merchants by Updates")
        top_merchants = filtered_df['Merchant'].value_counts().head(10)
        st.bar_chart(top_merchants)
    else:
        st.info("No data available for the selected date range.")

# ---------------------------
# TAB 2: Add Entry
# ---------------------------
with tab2:
    st.header("➕ Add Allocation Entry")

    # Single Entry
    st.subheader("Single Entry")
    merchant = st.selectbox("Select Existing Merchant", options=[""] + st.session_state.merchant_history)
    new_merchant = st.text_input("Or Enter New Merchant Name")
    final_merchant = new_merchant.strip() if new_merchant.strip() else merchant

    selected_couriers = st.multiselect("Select Courier(s)", COURIERS)
    remarks = st.text_input("Remarks (Optional)")

    if st.button("💾 Save Update"):
        if final_merchant and selected_couriers:
            new_entry = {
                "Date": datetime.now().strftime("%d-%b-%y"),
                "Merchant": final_merchant,
                "Courier": " | ".join(selected_couriers),
                "Remarks": remarks
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            df.to_excel(FILE_PATH, index=False, engine="openpyxl")
            if final_merchant not in st.session_state.merchant_history:
                st.session_state.merchant_history.append(final_merchant)
            st.success("✅ Update saved successfully!")
        else:
            st.warning("⚠️ Please enter Merchant Name and select at least one Courier.")

    # Batch Entry
    with st.expander("📝 Batch Entry (Multiple Merchants)"):
        batch_entries = st.text_area("Add Multiple Merchants (one per line)")
        batch_couriers = st.multiselect("Apply Couriers to All", COURIERS, key="batch_couriers_tab")
        if st.button("💾 Save Batch Entries"):
            added_count = 0
            if batch_entries and batch_couriers:
                for merchant_name in batch_entries.splitlines():
                    merchant_name = merchant_name.strip()
                    if merchant_name:
                        new_entry = {
                            "Date": datetime.now().strftime("%d-%b-%y"),
                            "Merchant": merchant_name,
                            "Courier": " | ".join(batch_couriers),
                            "Remarks": ""
                        }
                        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                        if merchant_name not in st.session_state.merchant_history:
                            st.session_state.merchant_history.append(merchant_name)
                        added_count += 1
                df.to_excel(FILE_PATH, index=False, engine="openpyxl")
                st.success(f"✅ {added_count} entries added!")

# ---------------------------
# TAB 3: Logs
# ---------------------------
with tab3:
    st.header("📄 Allocation Logs")
    search_merchant = st.text_input("Search by Merchant", key="search_tab3")
    search_courier = st.text_input("Search by Courier", key="search_courier_tab3")

    logs_df = df.copy()
    if search_merchant:
        logs_df = logs_df[logs_df["Merchant"].str.contains(search_merchant, case=False, na=False)]
    if search_courier:
        logs_df = logs_df[logs_df["Courier"].str.contains(search_courier, case=False, na=False)]

    # Format Date for display
    logs_df['Date'] = pd.to_datetime(logs_df['Date']).dt.strftime('%d-%b-%y')

    # Highlight last 7 days
    def highlight_recent(row):
        try:
            row_date = datetime.strptime(row['Date'], '%d-%b-%y').date()
            if row_date > datetime.now().date() - timedelta(days=7):
                return ['background-color: #d4f4dd']*len(row)
            else:
                return ['']*len(row)
        except:
            return ['']*len(row)

    st.dataframe(logs_df.tail(100).style.apply(highlight_recent, axis=1))

    # Delete Option
    with st.expander("🗑️ Delete Entry"):
        if not df.empty:
            options = [f"{i} | {row['Merchant']} | {row['Courier']}" for i, row in df.iterrows()]
            del_select = st.selectbox("Select entry to delete", options)
            confirm_del = st.checkbox("⚠️ Confirm deletion")
            if st.button("❌ Delete Selected Entry") and confirm_del:
                idx = int(del_select.split(" | ")[0])
                df = df.drop(index=idx).reset_index(drop=True)
                df.to_excel(FILE_PATH, index=False, engine="openpyxl")
                st.success("✅ Entry deleted successfully!")
        else:
            st.info("No entries to delete.")

    # Download Excel
    st.subheader("⬇️ Download Full Allocation Log")
    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine="openpyxl")
    towrite.seek(0)
    st.download_button("📥 Download Excel File", data=towrite,
                       file_name="allocation_updates.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
