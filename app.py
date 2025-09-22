# app.py
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

st.set_page_config(page_title="Allocation CRM", layout="wide")

# ---------------------------
# Config / Constants
# ---------------------------
DATA_FILE = "allocation_updates.xlsx"
RECYCLE_FILE = "recycle_bin.xlsx"
BACKUP_FOLDER = "backups"
os.makedirs(BACKUP_FOLDER, exist_ok=True)

# Fixed courier list
couriers = [
    "ATS", "NimbusPost", "Blue Dart Direct", "Shipyaari", "Delhivery",
    "Ecom Express", "Ekart", "Shadowfax", "Xpressbees", "Blitz",
    "Professional", "GoSwift", "Pikndel"
]

# Auto-suggest remarks
COMMON_REMARKS = [
    "Tested", "Pending Approval", "Issue Resolved", "Temporary Change",
    "Permanent Change", "Customer Request", "Operational Test"
]

# ---------------------------
# Helper functions
# ---------------------------
def load_data():
    if os.path.exists(DATA_FILE):
        return pd.read_excel(DATA_FILE)
    else:
        df0 = pd.DataFrame(columns=["Date", "Merchant", "Courier", "Remarks"])
        df0.to_excel(DATA_FILE, index=False)
        return df0

def load_recycle():
    if os.path.exists(RECYCLE_FILE):
        return pd.read_excel(RECYCLE_FILE)
    else:
        df0 = pd.DataFrame(columns=["Date", "Merchant", "Courier", "Remarks", "DeletedAt"])
        df0.to_excel(RECYCLE_FILE, index=False)
        return df0

def backup_df(df_local):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = os.path.join(BACKUP_FOLDER, f"allocation_backup_{ts}.xlsx")
    try:
        df_local.to_excel(fname, index=False)
    except Exception as e:
        st.warning("Backup failed: " + str(e))

def save_data(df_local):
    df_local.to_excel(DATA_FILE, index=False)
    backup_df(df_local)

def save_recycle(df_local):
    df_local.to_excel(RECYCLE_FILE, index=False)

def safe_to_datetime(s):
    try:
        return pd.to_datetime(s)
    except:
        return pd.NaT

def df_to_excel_bytes(df_in):
    towrite = io.BytesIO()
    df_in.to_excel(towrite, index=False, engine="openpyxl")
    towrite.seek(0)
    return towrite.getvalue()

def create_pdf_report(df_summary, top_merchants_df, out_path):
    # Creates a simple PDF with summary text and a table of top merchants
    try:
        with PdfPages(out_path) as pdf:
            # Page 1: Summary KPIs
            fig, ax = plt.subplots(figsize=(8.27, 11.69))  # A4 portrait
            ax.axis("off")
            lines = []
            lines.append("Allocation CRM - Summary Report")
            lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            lines.append("")
            for k, v in df_summary.items():
                lines.append(f"{k}: {v}")
            text = "\n".join(lines)
            ax.text(0.01, 0.98, text, fontsize=12, va='top')
            pdf.savefig(fig)
            plt.close(fig)

            # Page 2: Top Merchants table
            fig, ax = plt.subplots(figsize=(8.27, 11.69))
            ax.axis('off')
            ax.set_title("Top Merchants by Updates", fontsize=14, pad=20)
            # Draw table
            table = ax.table(cellText=top_merchants_df.values,
                             colLabels=top_merchants_df.columns,
                             loc='center',
                             cellLoc='left')
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)
            pdf.savefig(fig)
            plt.close(fig)
        return True, None
    except Exception as e:
        return False, str(e)

# ---------------------------
# Load initial data
# ---------------------------
df = load_data()
recycle_bin = load_recycle()

# Session-state for merchant history
if "merchant_history" not in st.session_state:
    st.session_state.merchant_history = df["Merchant"].dropna().unique().tolist()

# ---------------------------
# UI: Tabs (Dashboard, Add Entry, Logs)
# ---------------------------
tab1, tab2, tab3 = st.tabs(["📊 Dashboard", "➕ Add Entry", "📄 Logs"])

# ---------------------------
# TAB 1: Dashboard
# ---------------------------
with tab1:
    st.header("📊 Allocation Dashboard")

    # Date range filter for dashboard (defaults: last 30 days)
    col_dr1, col_dr2, col_dr3 = st.columns([1,1,2])
    with col_dr1:
        today = datetime.now().date()
        default_start = today - timedelta(days=30)
        date_range = st.date_input("Dashboard Date Range", value=(default_start, today))
    with col_dr2:
        granularity = st.selectbox("Granularity", ["Daily", "Weekly", "Monthly"])
    with col_dr3:
        st.markdown("**Quick Filters**")
        qmerchant = st.selectbox("Merchant (All = show all)", options=["All"] + st.session_state.merchant_history)
        qcourier = st.selectbox("Courier (All = show all)", options=["All"] + couriers)

    # Prepare filtered df for dashboard
    df_dash = df.copy()
    if len(date_range) == 2:
        start_d, end_d = date_range
        # ensure datetime comparison
        df_dash["Date_parsed"] = pd.to_datetime(df_dash["Date"], errors='coerce')
        df_dash = df_dash[(df_dash["Date_parsed"].dt.date >= start_d) & (df_dash["Date_parsed"].dt.date <= end_d)]

    if qmerchant != "All":
        df_dash = df_dash[df_dash["Merchant"] == qmerchant]
    if qcourier != "All":
        df_dash = df_dash[df_dash["Courier"].str.contains(qcourier, na=False)]

    # Summary cards
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Updates", len(df_dash))
    c2.metric("Unique Merchants", df_dash["Merchant"].nunique())
    c3.metric("Unique Couriers Used", df_dash["Courier"].nunique())
    # Most used courier within filtered set
    try:
        most_used = df_dash['Courier'].str.split(' \| ').explode().value_counts().idxmax()
    except Exception:
        most_used = "-"
    c4.metric("Most Used Courier", most_used)

    st.markdown("---")

    # Merchant Leaderboard (Top 10)
    with st.container():
        st.subheader("Merchant Leaderboard (Top 10)")
        top_merchants = df_dash['Merchant'].value_counts().head(10)
        if not top_merchants.empty:
            st.bar_chart(top_merchants)
        else:
            st.info("No merchant data for selected filters.")

    # Trendline per granularity
    with st.container():
        st.subheader(f"{granularity} Trend of Updates")
        if not df_dash.empty:
            df_dash["Date_parsed"] = pd.to_datetime(df_dash["Date"], errors='coerce')
            if granularity == "Daily":
                trend = df_dash.groupby(df_dash["Date_parsed"].dt.date).size()
            elif granularity == "Weekly":
                trend = df_dash.groupby(df_dash["Date_parsed"].dt.to_period("W")).size()
                trend.index = trend.index.astype(str)
            else:  # monthly
                trend = df_dash.groupby(df_dash["Date_parsed"].dt.to_period("M")).size()
                trend.index = trend.index.astype(str)
            st.line_chart(trend)
        else:
            st.info("No data to show for the selected timeframe.")

    # Merchant profile quick view
    with st.expander("🔎 Merchant Profile"):
        prof_merchant = st.selectbox("Select Merchant", options=[""] + st.session_state.merchant_history)
        if prof_merchant:
            prof_df = df[df["Merchant"] == prof_merchant].copy()
            if prof_df.empty:
                st.info("No records for this merchant.")
            else:
                st.write(f"Total records for **{prof_merchant}**: {len(prof_df)}")
                st.dataframe(prof_df.sort_values("Date", ascending=False).reset_index(drop=True))
                # Merchant trend
                prof_df["Date_parsed"] = pd.to_datetime(prof_df["Date"], errors='coerce')
                prof_trend = prof_df.groupby(prof_df["Date_parsed"].dt.date).size()
                st.line_chart(prof_trend)
                # Couriers used by this merchant
                courier_counts = prof_df['Courier'].str.split(' \| ').explode().value_counts()
                st.subheader("Couriers used by merchant")
                st.table(courier_counts.head(20))

    # Courier performance insights
    with st.expander("🚚 Courier Performance Insights"):
        st.subheader("Courier Usage (All time)")
        if not df.empty:
            courier_all = df['Courier'].str.split(' \| ').explode().value_counts()
            st.table(courier_all.head(20))
            st.write("Top 5 couriers:")
            st.bar_chart(courier_all.head(5))
        else:
            st.info("No data yet.")

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

    # Copy last settings button
    if final_merchant and st.button("⤷ Copy Last Allocation for Merchant"):
        last_row = df[df["Merchant"] == final_merchant].sort_values("Date", ascending=False)
        if not last_row.empty:
            last = last_row.iloc[0]
            # fill values in session state for convenience
            st.session_state._copy_couriers = last["Courier"].split(" | ")
            st.session_state._copy_remarks = last.get("Remarks", "")
            st.success("Copied last allocation into form. Reopen multiselect to see values.")
        else:
            st.info("No previous allocation found for this merchant.")

    # Provide selected couriers; if copy exists, preselect
    preselect = st.session_state.get("_copy_couriers", None)
    selected_couriers = st.multiselect("Select Courier(s)", couriers, default=preselect)
    remarks_default = st.session_state.get("_copy_remarks", "")
    remarks = st.selectbox("Quick Remarks (or type below)", options=[""] + COMMON_REMARKS, index=0)
    remarks_manual = st.text_input("Remarks (optional)", value=remarks if remarks else remarks_default)

    if st.button("💾 Save Update"):
        if not final_merchant:
            st.warning("Please enter or select a merchant.")
        elif not selected_couriers:
            st.warning("Please select at least one courier.")
        else:
            new_entry = {
                "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Merchant": final_merchant,
                "Courier": " | ".join(selected_couriers),
                "Remarks": remarks_manual
            }
            # Prevent duplicate same day same merchant+courier
            same = df[
                (df["Merchant"] == final_merchant) &
                (df["Courier"] == new_entry["Courier"]) &
                (pd.to_datetime(df["Date"]).dt.date == datetime.now().date())
            ]
            if not same.empty:
                st.warning("Similar entry already exists for this merchant today.")
            else:
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                save_data(df)
                # update session merchant history
                if final_merchant not in st.session_state.merchant_history:
                    st.session_state.merchant_history.append(final_merchant)
                # clear copy cache
                st.session_state.pop("_copy_couriers", None)
                st.session_state.pop("_copy_remarks", None)
                st.success("✅ Update saved successfully!")

    # Batch upload (file) + batch text option
    st.subheader("Bulk / Batch Upload")
    st.markdown("You can upload a CSV/XLSX with columns: Merchant, Courier (or multiple couriers separated by '|'), Remarks (optional).")
    uploaded_file = st.file_uploader("Upload CSV / Excel file", type=["csv", "xlsx"])
    if uploaded_file:
        try:
            if uploaded_file.name.lower().endswith(".csv"):
                batch_df = pd.read_csv(uploaded_file)
            else:
                batch_df = pd.read_excel(uploaded_file)
            st.write("Preview of uploaded file:")
            st.dataframe(batch_df.head())
            if st.button("💾 Save Uploaded Entries"):
                added = 0
                for _, row in batch_df.iterrows():
                    merchant_name = str(row.get("Merchant", "")).strip()
                    courier_val = str(row.get("Courier", "")).strip()
                    remarks_val = str(row.get("Remarks", "")).strip() if "Remarks" in row else ""
                    if merchant_name and courier_val:
                        new_entry = {
                            "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "Merchant": merchant_name,
                            "Courier": courier_val.replace(",", " | "),
                            "Remarks": remarks_val
                        }
                        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                        if merchant_name not in st.session_state.merchant_history:
                            st.session_state.merchant_history.append(merchant_name)
                        added += 1
                save_data(df)
                st.success(f"✅ {added} entries added from file.")
        except Exception as e:
            st.error("Failed to read uploaded file: " + str(e))

    # Text-based batch entry (one merchant per line)
    with st.expander("Enter multiple merchants (one per line)"):
        batch_text = st.text_area("Enter merchants here (one per line)")
        batch_couriers = st.multiselect("Apply couriers to all (if left empty, first column in CSV will be used)", couriers, key="batch_text")
        if st.button("💾 Save Batch Text Entries"):
            added = 0
            for line in batch_text.splitlines():
                m = line.strip()
                if m:
                    courier_val = " | ".join(batch_couriers) if batch_couriers else ""
                    new_entry = {
                        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "Merchant": m,
                        "Courier": courier_val,
                        "Remarks": ""
                    }
                    df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                    if m not in st.session_state.merchant_history:
                        st.session_state.merchant_history.append(m)
                    added += 1
            save_data(df)
            st.success(f"✅ {added} entries added.")

# ---------------------------
# TAB 3: Logs
# ---------------------------
with tab3:
    st.header("📄 Allocation Logs & Management")

    # Filters & search
    c1, c2, c3, c4 = st.columns([2,2,1,1])
    with c1:
        s_merchant = st.text_input("Search Merchant", value="")
    with c2:
        s_courier = st.text_input("Search Courier", value="")
    with c3:
        # date range filter
        today = datetime.now().date()
        dr = st.date_input("Date Range", value=(today - timedelta(days=90), today))
    with c4:
        show_deleted = st.checkbox("Show Recycle Bin", value=False)

    # choose which df to show
    if show_deleted:
        display_df = recycle_bin.copy()
        st.subheader("Recycle Bin (Deleted Entries)")
        if not display_df.empty:
            st.dataframe(display_df.sort_values("DeletedAt", ascending=False).reset_index(drop=True))
            # restore option
            with st.expander("Restore an entry"):
                idx_options = display_df.reset_index().apply(lambda r: f"{r['index']} | {r['Merchant']} | {r['DeletedAt']}", axis=1).tolist()
                # safer option: show merchant list
                merchant_to_restore = st.selectbox("Select merchant to restore", options=[""] + display_df["Merchant"].unique().tolist())
                if st.button("♻️ Restore Selected Entry"):
                    if merchant_to_restore:
                        to_restore = display_df[display_df["Merchant"] == merchant_to_restore].iloc[0]
                        restored_row = to_restore.drop(labels=["index", "DeletedAt"], errors='ignore')
                        # append to df
                        df = pd.concat([df, pd.DataFrame([restored_row[["Date","Merchant","Courier","Remarks"]].to_dict()])], ignore_index=True)
                        # remove from recycle bin
                        recycle_bin = recycle_bin[~((recycle_bin["Merchant"] == merchant_to_restore) & (recycle_bin["DeletedAt"] == to_restore["DeletedAt"]))]
                        save_data(df)
                        save_recycle(recycle_bin)
                        st.success("✅ Restored entry to main data.")
                    else:
                        st.warning("Please select a merchant to restore.")
        else:
            st.info("Recycle Bin empty.")
    else:
        # filter main df
        display_df = df.copy()
        # apply date range
        if len(dr) == 2:
            dstart, dend = dr
            display_df["Date_parsed"] = pd.to_datetime(display_df["Date"], errors='coerce')
            display_df = display_df[(display_df["Date_parsed"].dt.date >= dstart) & (display_df["Date_parsed"].dt.date <= dend)]
        if s_merchant:
            display_df = display_df[display_df["Merchant"].str.contains(s_merchant, case=False, na=False)]
        if s_courier:
            display_df = display_df[display_df["Courier"].str.contains(s_courier, case=False, na=False)]

        st.dataframe(display_df.sort_values("Date", ascending=False).reset_index(drop=True).tail(200))

        # Row-level delete (move to recycle bin)
        with st.expander("🗑️ Delete an Entry (moves to Recycle Bin)"):
            if not display_df.empty:
                # present choices as "index | merchant | courier | date"
                options = [
                    f"{i} | {r['Merchant']} | {r['Courier']} | {r['Date']}"
                    for i, r in display_df.reset_index(drop=False).iterrows()
                ]
                sel = st.selectbox("Select entry to delete", options=options)
                if sel:
                    if st.button("❌ Move to Recycle Bin"):
                        idx = int(sel.split(" | ")[0])
                        row = display_df.reset_index(drop=False).iloc[idx]
                        deleted_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        rec_row = {
                            "Date": row["Date"],
                            "Merchant": row["Merchant"],
                            "Courier": row["Courier"],
                            "Remarks": row.get("Remarks", ""),
                            "DeletedAt": deleted_at
                        }
                        recycle_bin = pd.concat([recycle_bin, pd.DataFrame([rec_row])], ignore_index=True)
                        # drop from main df: match by Date+Merchant+Courier (best effort)
                        mask = ~(
                            (df["Merchant"] == row["Merchant"]) &
                            (df["Courier"] == row["Courier"]) &
                            (df["Date"] == row["Date"])
                        )
                        df = df[mask].reset_index(drop=True)
                        save_data(df)
                        save_recycle(recycle_bin)
                        st.success("✅ Moved to Recycle Bin.")

        # Export options
        st.subheader("Export / Download")
        colx1, colx2 = st.columns(2)
        with colx1:
            if st.button("⬇️ Download Filtered Excel"):
                bytes_xlsx = df_to_excel_bytes(display_df)
                st.download_button("Download Excel", data=bytes_xlsx, file_name="allocation_filtered.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with colx2:
            if st.button("📄 Generate PDF Report (Summary + Top Merchants)"):
                # Prepare summary
                summary = {
                    "Total Updates (filtered)": len(display_df),
                    "Unique Merchants (filtered)": display_df["Merchant"].nunique(),
                    "Unique Couriers (filtered)": display_df["Courier"].nunique(),
                    "Report Generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                topm = display_df['Merchant'].value_counts().reset_index()
                topm.columns = ["Merchant", "Updates"]
                if topm.empty:
                    st.warning("No data to include in PDF.")
                else:
                    # create temp pdf file
                    pdf_name = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                    success, err = create_pdf_report(summary, topm.head(20), pdf_name)
                    if success:
                        with open(pdf_name, "rb") as f:
                            pdf_bytes = f.read()
                        st.download_button("Download PDF Report", data=pdf_bytes, file_name=pdf_name, mime="application/pdf")
                    else:
                        st.error("Failed to create PDF: " + str(err))

# ---------------------------
# Final: Persist global df & recycle into session scope (optional)
# ---------------------------
# Write final df back to disk in case of last operations outside Save button
try:
    save_data(df)
    save_recycle(recycle_bin)
except Exception:
    pass

st.sidebar.markdown("---")
st.sidebar.info("Allocation CRM — Final upgraded version.\nBackups saved in `backups/`.")
