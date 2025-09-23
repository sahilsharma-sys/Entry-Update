import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io

# 📌 Fixed courier list
couriers = [
    "ATS", "NimbusPost", "Blue Dart Direct", "Shipyaari", "Delhivery",
    "Ecom Express", "Ekart", "Shadowfax", "Xpressbees", "Blitz",
    "Professional", "GoSwift", "Pikndel", "DTDC"
]

# 📌 File path
file_path = "allocation_updates.xlsx"

# 🔹 Ensure file exists
if not os.path.exists(file_path):
    pd.DataFrame(columns=["Date", "Merchant", "Courier", "Remarks"]).to_excel(file_path, index=False, engine="openpyxl")

# --- Session State for Merchant History ---
if "merchant_history" not in st.session_state:
    df_init = pd.read_excel(file_path)
    st.session_state.merchant_history = df_init["Merchant"].dropna().unique().tolist()

# --- CRM-Style Tabs ---
tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "➕ Add Entry", "📄 Logs"])

# ----------------- TAB 1: Dashboard -----------------
with tab1:
    st.header("📊 Allocation Dashboard")

    df = pd.read_excel(file_path)   # always load fresh

    if not df.empty:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

        # Date Filter
        min_date, max_date = df["Date"].min(), df["Date"].max()
        date_range = st.date_input("📅 Filter by Date Range", [min_date, max_date])
        filtered_df = df.copy()
        if isinstance(date_range, list) and len(date_range) == 2:
            filtered_df = filtered_df[(filtered_df["Date"] >= pd.to_datetime(date_range[0])) & (filtered_df["Date"] <= pd.to_datetime(date_range[1]))]

        # Summary Cards
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Updates", len(filtered_df))
        col2.metric("Unique Merchants", filtered_df["Merchant"].nunique())
        col3.metric("Unique Couriers Used", filtered_df["Courier"].nunique())

        # Last Update Info
        last_update = df["Date"].max().strftime("%Y-%m-%d %H:%M:%S")
        st.info(f"🕒 Last Update: {last_update}")

        # Insights
        top_courier = filtered_df['Courier'].str.split(' \| ').explode().value_counts().idxmax()
        top_merchant = filtered_df['Merchant'].value_counts().idxmax()
        col1, col2 = st.columns(2)
        col1.metric("Most Used Courier", top_courier)
        col2.metric("Most Updated Merchant", top_merchant)

        # Weekly Trend
        st.subheader("📈 Updates Over Time (Weekly)")
        weekly = filtered_df.set_index("Date").resample("W").size()
        st.line_chart(weekly)

        # Top Merchants
        st.subheader("🛍️ Top 10 Merchants by Updates")
        top_merchants = filtered_df['Merchant'].value_counts().head(10)
        st.bar_chart(top_merchants)
    else:
        st.info("No data available yet.")

# ----------------- TAB 2: Add Entry -----------------
with tab2:
    st.header("➕ Add Allocation Entry")

    df = pd.read_excel(file_path)   # reload

    # Single Entry
    st.subheader("Single Entry")
    merchant = st.selectbox(
        "Select Existing Merchant",
        options=[""] + st.session_state.merchant_history,
        index=0
    )
    new_merchant = st.text_input("Or Enter New Merchant Name")
    final_merchant = new_merchant if new_merchant else merchant

    selected_couriers = st.multiselect("Select Courier(s)", couriers)
    replace_with = st.text_input("Remarks (Optional)")

    if st.button("💾 Save Update"):
        if final_merchant and selected_couriers:
            new_entry = {
                "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Merchant": final_merchant,
                "Courier": " | ".join(selected_couriers),
                "Remarks": replace_with if replace_with else ""
            }
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            df.to_excel(file_path, index=False, engine="openpyxl")

            if final_merchant not in st.session_state.merchant_history:
                st.session_state.merchant_history.append(final_merchant)

            st.success("✅ Update saved successfully! Please refresh Dashboard/Logs to see.")
        else:
            st.warning("⚠️ Please enter Merchant Name and select at least one Courier.")

    # Batch Entry
    with st.expander("📝 Batch Entry (Multiple Merchants)"):

        batch_entries = st.text_area("Add Multiple Merchants (one per line)")
        batch_couriers = st.multiselect("Apply Couriers to All", couriers, key="batch_couriers_tab")
        if st.button("💾 Save Batch Entries"):
            if batch_entries and batch_couriers:
                added_count = 0
                for merchant_name in batch_entries.split("\n"):
                    merchant_name = merchant_name.strip()
                    if merchant_name:
                        new_entry = {
                            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Merchant": merchant_name,
                            "Courier": " | ".join(batch_couriers),
                            "Remarks": ""
                        }
                        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                        if merchant_name not in st.session_state.merchant_history:
                            st.session_state.merchant_history.append(merchant_name)
                        added_count += 1
                df.to_excel(file_path, index=False, engine="openpyxl")
                st.success(f"✅ {added_count} entries added! Refresh Dashboard/Logs to see.")

# ----------------- TAB 3: Logs -----------------
with tab3:
    st.header("📄 Allocation Logs")

    df = pd.read_excel(file_path)   # reload

    # Search / Filter
    search_merchant = st.text_input("Search by Merchant", key="search_tab3")
    search_courier = st.text_input("Search by Courier", key="search_courier_tab3")

    filtered_df = df.copy()
    if search_merchant:
        filtered_df = filtered_df[filtered_df["Merchant"].str.contains(search_merchant, case=False, na=False)]
    if search_courier:
        filtered_df = filtered_df[filtered_df["Courier"].str.contains(search_courier, case=False, na=False)]

    # Highlight Recent Entries
    def highlight_recent(row):
        try:
            if pd.to_datetime(row['Date']) > datetime.now() - timedelta(days=7):
                return ['background-color: #d4f4dd'] * len(row)
            else:
                return [''] * len(row)
        except:
            return [''] * len(row)

    st.dataframe(filtered_df.tail(50).style.apply(highlight_recent, axis=1))

    # Delete Option
    with st.expander("🗑️ Delete Entry"):
        if not df.empty:
            delete_option = st.selectbox(
                "Select entry to delete",
                [f"{i} | {row['Merchant']} | {row['Courier']}" for i, row in df.iterrows()]
            )
            confirm_delete = st.checkbox("⚠️ Confirm deletion")
            if st.button("❌ Delete Selected Entry") and confirm_delete:
                row_index = int(delete_option.split(" | ")[0])
                df = df.drop(index=row_index).reset_index(drop=True)
                df.to_excel(file_path, index=False, engine="openpyxl")
                st.success("✅ Entry deleted successfully! Refresh Logs to see.")
        else:
            st.info("No entries available to delete.")

    # Download
    st.subheader("⬇️ Download Allocation Log")
    towrite = io.BytesIO()
    df.to_excel(towrite, index=False, engine="openpyxl")
    towrite.seek(0)
    st.download_button(
        label="📥 Download Excel File",
        data=towrite,
        file_name="allocation_updates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
