import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
import io, zipfile, tempfile, os

st.set_page_config(page_title="Master Excel Utility Tool", layout="wide")

# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.title("🛠 Master Utility Tool")
menu = st.sidebar.radio(
    "Select Tool",
    [
        "📂 New File Creation",
        "🔄 CSV → XLSX Converter",
        "📝 Merchant Auto Rename",
    ]
)

st.sidebar.info(
    "☁ Cloud Version\n"
    "📤 Upload Files / Folder (ZIP)\n"
    "📥 Download Output (ZIP)\n\n"
    "❌ No path paste required"
)

# =====================================================
# HELPERS
# =====================================================
def extract_files(files, zip_file, exts):
    extracted = []

    if files:
        extracted.extend(files)

    if zip_file:
        tmp = tempfile.mkdtemp()
        zip_path = os.path.join(tmp, zip_file.name)
        with open(zip_path, "wb") as f:
            f.write(zip_file.read())

        with zipfile.ZipFile(zip_path) as z:
            z.extractall(tmp)

        for root, _, filenames in os.walk(tmp):
            for name in filenames:
                if name.lower().endswith(exts):
                    extracted.append(open(os.path.join(root, name), "rb"))
    return extracted


def make_zip(file_buffers):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in file_buffers:
            z.writestr(name, data.getvalue())
    buf.seek(0)
    return buf


# =====================================================
# 2️⃣ CSV → XLSX
# =====================================================
elif menu == "🔄 CSV → XLSX Converter":
    st.title("🔄 CSV → XLSX Converter")

    files = st.file_uploader("📤 Upload CSV Files", type="csv", accept_multiple_files=True)
    zip_file = st.file_uploader("📦 OR Upload CSV Folder (ZIP)", type="zip")

    if st.button("🚀 Convert"):
        csvs = extract_files(files, zip_file, (".csv",))
        if not csvs:
            st.error("❌ No CSV files")
            st.stop()

        output = []
        for f in csvs:
            df = pd.read_csv(f)
            buf = io.BytesIO()
            df.to_excel(buf, index=False)
            buf.seek(0)
            output.append((f.name.replace(".csv", ".xlsx"), buf))

        st.download_button("📥 Download XLSX ZIP", make_zip(output), "Converted_XLSX.zip", "application/zip")

# =====================================================
# 3️⃣ MERCHANT AUTO RENAME
# =====================================================
elif menu == "📝 Merchant Auto Rename":
    st.title("📝 Merchant Auto Rename")

    files = st.file_uploader("📤 Upload Excel Files", type="xlsx", accept_multiple_files=True)
    zip_file = st.file_uploader("📦 OR Upload Folder (ZIP)", type="zip")

    if st.button("🚀 Rename"):
        excels = extract_files(files, zip_file, (".xlsx",))
        output = []

        for f in excels:
            df = pd.read_excel(f, header=None)
            name = f.name
            if str(df.iloc[0,0]).lower() == "client name":
                name = str(df.iloc[1,0]).strip() + ".xlsx"

            buf = io.BytesIO()
            df.to_excel(buf, index=False, header=False)
            buf.seek(0)
            output.append((name, buf))

        st.download_button("📥 Download Renamed ZIP", make_zip(output), "Renamed_Files.zip", "application/zip")
