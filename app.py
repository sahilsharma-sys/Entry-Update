import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
import io, zipfile, tempfile, os

st.set_page_config(page_title="Master Excel Utility Tool", layout="wide")

# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.title("🛠 Master Excel Utility Tool")
menu = st.sidebar.radio(
    "Select Tool",
    [
        "📂 New File Creation",
        "📊 Excel Formula Updater"
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
def upload_block(title, types):
    st.subheader(f"📁 {title}")
    files = st.file_uploader(
        f"Upload {title} (Files)",
        type=types,
        accept_multiple_files=True,
        key=title+"_files"
    )
    zip_file = st.file_uploader(
        f"Upload {title} (Folder ZIP)",
        type="zip",
        key=title+"_zip"
    )
    return files, zip_file


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
# 📂 NEW FILE CREATION (MULTI-PATH STYLE)
# =====================================================
if menu == "📂 New File Creation":
    st.title("📂 New File Creation Tool")

    st.markdown("### 🔹 Reference Files")
    adf_file = st.file_uploader("📘 Upload ADF Foods File", type="xlsx")
    pricing_file = st.file_uploader("📗 Upload Pricing Format File", type="xlsx")

    st.markdown("---")
    raw_files, raw_zip = upload_block("RAW FILES", ["xlsx"])
    error_files, error_zip = upload_block("ERROR FILES", ["xlsx"])

    if st.button("🚀 Start Processing"):
        if not adf_file or not pricing_file:
            st.error("❌ Upload reference files")
            st.stop()

        raw_excels = extract_files(raw_files, raw_zip, (".xlsx",))
        error_excels = extract_files(error_files, error_zip, (".xlsx",))

        all_files = raw_excels + error_excels
        if not all_files:
            st.error("❌ No Excel files uploaded")
            st.stop()

        ref_wb = load_workbook(adf_file, data_only=False)
        ref_ws = ref_wb.active
        ref_formulas = {
            c: ref_ws.cell(2, c).value
            for c in range(38, 65)
            if isinstance(ref_ws.cell(2, c).value, str)
            and ref_ws.cell(2, c).value.startswith("=")
        }

        pricing_df = pd.read_excel(pricing_file, header=None)
        pricing_data = pricing_df.values.tolist()

        extra_cols = [
            "Weight","Merchant Zone","Courier Name","FWD","CHECK","COD","CHECK",
            "RTO","CHECK","REVERSAL","CHECK","QC","TOTAL","TOTAL+GST","Diff",
            "Courier charged weight","Courier Zone","Freight","RTO","RTO Discount",
            "Reverse","COD","SDL","Fuel","QC","Others","Gross"
        ]

        output = []
        prog = st.progress(0)

        for i, f in enumerate(all_files, start=1):
            df = pd.read_excel(f, usecols=range(37))
            wb = load_workbook(f)
            ws = wb.active
            ws.title = "Raw"

            ws.delete_rows(1, ws.max_row)
            ws.append(list(df.columns))
            for _, row in df.iterrows():
                ws.append(list(row))

            for j, col in enumerate(extra_cols):
                ws.cell(1, 38 + j).value = col

            last_row = ws.max_row
            for col, formula in ref_formulas.items():
                letter = ws.cell(2, col).column_letter
                for r in range(2, last_row + 1):
                    ws[f"{letter}{r}"] = Translator(
                        formula, f"{letter}2"
                    ).translate_formula(f"{letter}{r}")

            if "Pricing" in wb.sheetnames:
                wb.remove(wb["Pricing"])
            ps = wb.create_sheet("Pricing")
            for r, row in enumerate(pricing_data, start=1):
                for c, v in enumerate(row, start=1):
                    ps.cell(r, c, v)

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            output.append((f.name, buf))
            prog.progress(i / len(all_files))

        st.success("🎉 Files processed successfully")
        st.download_button(
            "📥 Download Output ZIP",
            make_zip(output),
            "New_Files_Output.zip",
            "application/zip"
        )

# =====================================================
# 📊 EXCEL FORMULA UPDATER (SEPARATE PATH STYLE)
# =====================================================
elif menu == "📊 Excel Formula Updater":
    st.title("📊 Excel Formula Updater")

    input_files, input_zip = upload_block("INPUT FILES", ["xlsx"])

    if st.button("🚀 Update Formula"):
        excels = extract_files(input_files, input_zip, (".xlsx",))
        if not excels:
            st.error("❌ No Excel files uploaded")
            st.stop()

        output = []
        for f in excels:
            wb = load_workbook(f)
            ws = wb.active

            for r in range(2, ws.max_row + 1):
                ws[f"AL{r}"] = f"=IF(W{r}/1000>9.9,CEILING(W{r}/1000,1),CEILING(W{r}/1000,0.5))"

            buf = io.BytesIO()
            wb.save(buf)
            buf.seek(0)
            output.append((f.name, buf))

        st.success("✅ Formula updated")
        st.download_button(
            "📥 Download Updated ZIP",
            make_zip(output),
            "Formula_Updated.zip",
            "application/zip"
        )
