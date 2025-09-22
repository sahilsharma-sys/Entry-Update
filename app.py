import streamlit as st
import pandas as pd
import os
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# 📌 Fixed courier list
couriers = [
    "ATS", "NimbusPost", "Blue Dart Direct", "Shipyaari", "Delhivery",
    "Ecom Express", "Ekart", "Shadowfax"
]

# 📂 Excel file for storage
file_path = "merchant_logs.xlsx"

# Agar file exist karti hai to load karo
if os.path.exists(file_path):
    df = pd.read_excel(file_path)
else:
    df = pd.DataFrame(columns=["Merchant", "Courier", "Date"])

st.title("📦 CRM Allocation Update Tool")

# ✅ Merchant name input with suggestions
merchant = st.text_input("Enter Merchant Name")
courier = st.selectbox("Select Courier", couriers)

# Save Entry
if st.button("Add Entry"):
    if merchant.strip() != "":
        new_entry = {"Merchant": merchant, "Courier": courier, "Date": pd.Timestamp.now()}
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(file_path, index=False, engine="openpyxl")
        st.success(f"Entry saved for {merchant} → {courier}")

# 📊 Show Logs
st.subheader("📜 Logs")
st.dataframe(df)

# 🔎 Search & Filter
search_query = st.text_input("🔍 Search Merchant")
if search_query:
    filtered_df = df[df["Merchant"].str.contains(search_query, case=False, na=False)]
else:
    filtered_df = df

st.dataframe(filtered_df)

# 🗑 Delete Merchant Entry
delete_merchant = st.selectbox("Select Merchant to Delete", [""] + df["Merchant"].unique().tolist())
if st.button("Delete Entry"):
    if delete_merchant:
        df = df[df["Merchant"] != delete_merchant]
        df.to_excel(file_path, index=False, engine="openpyxl")
        st.warning(f"Deleted all entries for {delete_merchant}")

# 📥 Download as Excel
def convert_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Logs")
    processed_data = output.getvalue()
    return processed_data

st.download_button(
    label="📥 Download Logs (Excel)",
    data=convert_excel(df),
    file_name="merchant_logs.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# 📄 Download as PDF
def convert_pdf(df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 30, "Merchant Allocation Logs")

    c.setFont("Helvetica", 10)
    y = height - 60
    for i, row in df.iterrows():
        text = f"{row['Date']} | {row['Merchant']} → {row['Courier']}"
        c.drawString(30, y, text)
        y -= 15
        if y < 40:  # next page
            c.showPage()
            y = height - 40
            c.setFont("Helvetica", 10)

    c.save()
    buffer.seek(0)
    return buffer

st.download_button(
    label="📄 Download Logs (PDF)",
    data=convert_pdf(df),
    file_name="merchant_logs.pdf",
    mime="application/pdf"
)

# 📊 Dashboard Summary
st.subheader("📈 Summary Dashboard")
if not df.empty:
    col1, col2 = st.columns(2)
    with col1:
        st.metric("Total Entries", len(df))
        st.metric("Unique Merchants", df["Merchant"].nunique())
    with col2:
        st.metric("Unique Couriers", df["Courier"].nunique())
        top_merchant = df["Merchant"].value_counts().idxmax()
        st.metric("Top Merchant", top_merchant)
