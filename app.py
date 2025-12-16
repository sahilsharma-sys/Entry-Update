
import streamlit as st
import os
import shutil
import pandas as pd
import io
import openpyxl
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
import re

st.set_page_config(page_title="Master Excel Utility Tool", layout="wide")

# =====================================================
# SIDEBAR MENU
# =====================================================
st.sidebar.title("🛠 Master Utility Tool")
menu = st.sidebar.radio(
    "Select Tool",
    [
        "📂 New File Creation",
        "🔄 CSV → XLSX Converter",
        "📝 Merchant Auto-Rename Tool",
        "🗑 Delete Files Tool",
        "📊 Excel Formula Updater Tool",
        "🚚 Courier Cost Updater"
    ]
)

def live_file_status(folder_path, extra_files=[]):
    folder_path = os.path.normpath(folder_path.strip())
    if os.path.exists(folder_path):
        all_files = [f for f in os.listdir(folder_path) if f.lower().endswith((".xlsx", ".csv", ".xls"))]
    else:
        all_files = []

    all_files += extra_files

    st.subheader("📌 Live File Status")
    col1, col2 = st.columns(2)
    col1.metric("📁 Total Files", len(all_files))
    col2.metric("⏳ Pending Files", len(extra_files))

    st.expander("📂 Files").write(pd.DataFrame(all_files, columns=["Files"]))
    return all_files

# =====================================================
# 1️⃣ NEW FILE CREATION TOOL
# =====================================================
if menu == "📂 New File Creation":
    st.title("📊 New File Creation Tool")

    source_folder = st.text_input("📂 Source Folder Path", r"D:\Sahil\Invoices\Python\New file to create")
    working_folder = st.text_input("📁 Destination Folder Path", r"D:\Sahil\Invoices\Python\New file to create\New files")
    adf_food_file_path = st.text_input("📘 ADF Foods File Path", r"D:\Sahil\Invoices\Python\Working files\Adf Foods.xlsx")
    pricing_file_path = st.text_input("📗 Pricing Format File Path", r"D:\Sahil\Invoices\Python\Pricing Format\Pricing format.xlsx")

    source_files = live_file_status(source_folder)
    start_btn = st.button("🚀 Start Processing")

    extra_columns = [
        "Weight","Merchant Zone","Courier Name","FWD","CHECK","COD","CHECK",
        "RTO","CHECK","REVERSAL","CHECK","QC","TOTAL","TOTAL+GST","Diff",
        "Courier charged weight","Courier Zone","Freight","RTO","RTO Discount",
        "Reverse","COD","SDL","Fuel","QC","Others","Gross"
    ]

    if start_btn:
        ref_wb = load_workbook(adf_food_file_path, data_only=False)
        ref_ws = ref_wb.active
        ref_formulas = {col: ref_ws.cell(2, col).value for col in range(38,65) if isinstance(ref_ws.cell(2, col).value,str) and ref_ws.cell(2, col).value.startswith("=")}
        pricing_df = pd.read_excel(pricing_file_path, header=None)
        pricing_data = pricing_df.values.tolist()

        processed = []
        progress = st.progress(0)
        logs = st.empty()

        for i, filename in enumerate(source_files, start=1):
            try:
                logs.write(f"🔄 ({i}/{len(source_files)}) Processing: {filename}")
                src = os.path.join(source_folder, filename)
                dst = os.path.join(working_folder, filename)

                if not os.path.exists(working_folder):
                    os.makedirs(working_folder)

                shutil.copy2(src, dst)

                df = pd.read_excel(src, usecols=range(37))
                wb = load_workbook(dst, keep_links=False)

                ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
                if ws.title == "Sheet1":
                    ws.title = "Raw"

                for r in range(1, ws.max_row + 1):
                    for c in range(1, 38):
                        ws.cell(r, c).value = None

                for c, col_name in enumerate(df.columns, start=1):
                    ws.cell(1, c).value = col_name

                for ridx, row in df.iterrows():
                    for c, val in enumerate(row, start=1):
                        ws.cell(ridx + 2, c).value = val

                last_row = 2
                while ws.cell(last_row, 1).value:
                    last_row += 1
                last_row -= 1

                for j, name in enumerate(extra_columns):
                    ws.cell(1, 38 + j).value = name

                for col, base_formula in ref_formulas.items():
                    col_letter = ws.cell(2, col).column_letter
                    for r in range(2, last_row + 1):
                        ws[f"{col_letter}{r}"] = Translator(base_formula, origin=f"{col_letter}2").translate_formula(f"{col_letter}{r}")

                if "Pricing" in wb.sheetnames:
                    wb.remove(wb["Pricing"])
                pws = wb.create_sheet("Pricing")
                for r, row in enumerate(pricing_data, start=1):
                    for c, val in enumerate(row, start=1):
                        pws.cell(r, c, val)

                wb.save(dst)
                processed.append(filename)
                logs.write(f"✅ Processed: {filename}")

            except Exception as e:
                logs.write(f"❌ Error in {filename} → {e}")

            progress.progress(i / len(source_files))

        st.success("🎉 Processing Completed!")
        st.dataframe(pd.DataFrame(processed, columns=["Processed Files"]))

# =====================================================
# 2️⃣ CSV → XLSX CONVERTER
# =====================================================
elif menu == "🔄 CSV → XLSX Converter":
    st.title("🔄 CSV → XLSX Converter + Delete Tool")

    folder_path = st.text_input("📂 Folder Path", r"D:\Sahil\Invoices\Python\Monthly Data files")
    source_files = live_file_status(folder_path)

    option = st.radio("Choose Action", ["Convert CSV → XLSX", "Delete CSV Files", "Delete XLSX Files"], horizontal=True)

    if option == "Convert CSV → XLSX":
        convert_btn = st.button("🚀 Convert Now")
        if convert_btn:
            progress = st.progress(0)
            logs = st.empty()
            done = []
            csv_files = [f for f in source_files if f.lower().endswith(".csv")]

            for i, filename in enumerate(csv_files, start=1):
                try:
                    df = pd.read_csv(os.path.join(folder_path, filename))
                    df.to_excel(os.path.join(folder_path, filename[:-4] + ".xlsx"), index=False)
                    done.append(filename)
                    logs.write(f"✅ Converted: {filename}")
                except Exception as e:
                    logs.write(f"❌ Failed: {filename} → {e}")

                progress.progress(i / len(csv_files))

            st.success("🎉 Conversion Completed!")

    elif option == "Delete CSV Files":
        if st.button("🗑 Delete All CSV"):
            csv_files = [f for f in source_files if f.lower().endswith(".csv")]
            for f in csv_files:
                os.remove(os.path.join(folder_path, f))
            st.success("🗑 CSV Files Deleted!")

    elif option == "Delete XLSX Files":
        if st.button("🗑 Delete All XLSX"):
            xlsx_files = [f for f in source_files if f.lower().endswith(".xlsx")]
            for f in xlsx_files:
                os.remove(os.path.join(folder_path, f))
            st.success("🗑 XLSX Files Deleted!")

# =====================================================
# 3️⃣ MERCHANT AUTO-RENAME TOOL
# =====================================================
elif menu == "📝 Merchant Auto-Rename Tool":
    st.title("📝 Merchant Auto-Rename Tool (Client Name → File Name)")
    folder_path = st.text_input("📁 Folder Path")
    source_files = live_file_status(folder_path)

    if st.button("🚀 Start Renaming"):
        logs = st.empty()
        progress = st.progress(0)
        renamed, skipped, errors = [], [], []

        for i, filename in enumerate(source_files, start=1):
            try:
                file_path = os.path.join(folder_path, filename)
                df = pd.read_excel(file_path, header=None)
                if str(df.iloc[0,0]).lower() == "client name":
                    new_name = str(df.iloc[1,0]).strip() + ".xlsx"
                    new_name = "".join(x for x in new_name if x not in '<>:"/\\|?*')
                    os.rename(file_path, os.path.join(folder_path, new_name))
                    renamed.append(f"{filename} → {new_name}")
                else:
                    skipped.append(filename)
            except Exception as e:
                errors.append(f"{filename} → {e}")

            progress.progress(i / len(source_files))

        st.dataframe(pd.DataFrame(renamed, columns=["Renamed Files"]))

# =====================================================
# 4️⃣ DELETE FILES TOOL
# =====================================================
elif menu == "🗑 Delete Files Tool":
    st.title("🗑 Delete Files Based on Excel List")

    target_folder = st.text_input("📁 Target Folder Path")
    delete_file = st.text_input("📄 Delete List File Path")
    source_files = live_file_status(target_folder)

    if st.button("🚀 Delete Files"):
        df = pd.read_excel(delete_file)
        names = df.iloc[:,0].dropna().astype(str).tolist()
        deleted = []
        not_found = []

        for nm in names:
            found = False
            for f in source_files:
                if f.lower().startswith(nm.lower()):
                    os.remove(os.path.join(target_folder, f))
                    deleted.append(f)
                    found = True
                    break
            if not found:
                not_found.append(nm)

        st.dataframe(pd.DataFrame(deleted, columns=["Deleted Files"]))

# =====================================================
# 5️⃣ EXCEL FORMULA UPDATER TOOL (WITH NEW FILE CREATOR & SUMMARY DOWNLOAD)
# =====================================================
elif menu == "📊 Excel Formula Updater Tool":

    st.title("📊 Excel Formula Updater Tool (with Error → New Files Creator)")

    source_folder = st.text_input("📂 Source Folder Path")
    working_folder = st.text_input("📁 Destination Folder Path")

    adf_food_file_path = st.text_input(
        "📘 ADF Foods File Path",
        r"D:\Sahil\Invoices\Python\Working files\Adf Foods.xlsx"
    )
    pricing_file_path = st.text_input(
        "📗 Pricing Format File Path",
        r"D:\Sahil\Invoices\Python\Pricing Format\Pricing format.xlsx"
    )

    source_files = live_file_status(source_folder)

    # ----------------------------
    # Load reference formulas
    # ----------------------------
    ref_wb = load_workbook(adf_food_file_path, data_only=False)
    ref_ws = ref_wb.active

    ref_formulas = {
        col: ref_ws.cell(2, col).value
        for col in range(38, 65)
        if isinstance(ref_ws.cell(2, col).value, str)
        and ref_ws.cell(2, col).value.startswith("=")
    }

    # ----------------------------
    # Load pricing data
    # ----------------------------
    pricing_df = pd.read_excel(pricing_file_path, header=None)
    pricing_data = pricing_df.values.tolist()

    # Buttons
    start_btn = st.button("🚀 Start Updating")
    create_btn = st.button("🆕 Create New Files for Error Files")

    # =====================================================
    # START UPDATING
    # =====================================================
    if start_btn:

        progress = st.progress(0)
        logs = st.empty()

        processed = []
        error_files = []

        col1, col2, col3 = st.columns(3)
        col1.metric("✅ Processed", 0)
        col2.metric("❌ Errors", 0)
        col3.metric("⏳ Remaining", len(source_files))

        total_files = len(source_files)

        for i, filename in enumerate(source_files, start=1):
            try:
                source_file = os.path.join(source_folder, filename)
                working_file = os.path.join(working_folder, filename)

                df = pd.read_excel(source_file, usecols=range(37))
                wb = load_workbook(working_file)

                ws = None
                for sheet in wb.sheetnames:
                    if wb[sheet]["A1"].value == "Client Name":
                        ws = wb[sheet]
                        break
                if ws is None:
                    ws = wb.create_sheet("Sheet1")

                # Clear old data
                for r in range(1, ws.max_row + 1):
                    for c in range(1, 38):
                        ws.cell(r, c).value = None

                # Write headers
                for c, name in enumerate(df.columns, start=1):
                    ws.cell(1, c).value = name

                # Write data
                for ridx, row in df.iterrows():
                    for c, val in enumerate(row, start=1):
                        ws.cell(ridx + 2, c).value = val

                # AL column formula
                last_row = ws.max_row
                for r in range(2, last_row + 1):
                    ws[f"AL{r}"] = (
                        f"=IF(W{r}/1000>9.9,"
                        f"CEILING(W{r}/1000,1),"
                        f"CEILING(W{r}/1000,0.5))"
                    )

                wb.save(working_file)
                processed.append(filename)

                # DELETE source file after success
                os.remove(source_file)

                logs.write(f"✅ ({i}/{total_files}) Processed & Deleted: {filename}")

            except Exception as e:
                error_files.append(filename)
                logs.write(f"❌ ({i}/{total_files}) Error in {filename} → {e}")

            progress.progress(i / total_files)

            col1.metric("✅ Processed", len(processed))
            col2.metric("❌ Errors", len(error_files))
            col3.metric("⏳ Remaining", total_files - len(processed) - len(error_files))

        st.success("🎉 Update Completed!")

        st.write("✅ Processed Files")
        st.dataframe(pd.DataFrame(processed, columns=["Processed"]))

        st.write("❌ Error Files")
        st.dataframe(pd.DataFrame(error_files, columns=["Errors"]))

        # Save session state
        st.session_state["error_files"] = error_files
        st.session_state["source_folder"] = source_folder
        st.session_state["working_folder"] = working_folder

    # =====================================================
# 🆕 NEW FILE CREATOR FOR ERROR FILES
# (NEW FILES WILL BE CREATED INSIDE SOURCE FOLDER)
# =====================================================
if create_btn:

    if "error_files" not in st.session_state:
        st.error("❌ First run the updater to detect error files!")
    else:
        st.info("⚙ Creating NEW files inside SOURCE folder (New Files)…")

        error_files = st.session_state["error_files"]
        source_folder = st.session_state["source_folder"]
        working_folder = st.session_state["working_folder"]

        # 🔥 SAME LOGIC AS "NEW FILE CREATION TOOL"
        new_folder = os.path.join(source_folder, "New Files")
        os.makedirs(new_folder, exist_ok=True)

        created = []

        extra_columns = [
            "Weight","Merchant Zone","Courier Name","FWD","CHECK","COD","CHECK",
            "RTO","CHECK","REVERSAL","CHECK","QC","TOTAL","TOTAL+GST","Diff",
            "Courier charged weight","Courier Zone","Freight","RTO","RTO Discount",
            "Reverse","COD","SDL","Fuel","QC","Others","Gross"
        ]

        for f in error_files:
            try:
                # ✅ base file working folder se uthao
                src = os.path.join(source_folder, f)
                dst = os.path.join(new_folder, f)

                if not os.path.exists(src):
                    st.warning(f"⚠ File not found: {f}")
                    continue

                shutil.copy2(src, dst)

                df = pd.read_excel(dst, usecols=range(37))
                wb = load_workbook(dst, keep_links=False)

                ws = wb["Sheet1"] if "Sheet1" in wb.sheetnames else wb.active
                ws.title = "Raw"

                # Clear first 37 columns
                for r in range(1, ws.max_row + 1):
                    for c in range(1, 38):
                        ws.cell(r, c).value = None

                # Write headers
                for c, col_name in enumerate(df.columns, start=1):
                    ws.cell(1, c).value = col_name

                # Write data
                for ridx, row in df.iterrows():
                    for c, val in enumerate(row, start=1):
                        ws.cell(ridx + 2, c).value = val

                last_row = ws.max_row

                # Extra columns
                for j, name in enumerate(extra_columns):
                    ws.cell(1, 38 + j).value = name

                # Apply formulas
                for col, base_formula in ref_formulas.items():
                    col_letter = ws.cell(2, col).column_letter
                    for r in range(2, last_row + 1):
                        ws[f"{col_letter}{r}"] = Translator(
                            base_formula,
                            origin=f"{col_letter}2"
                        ).translate_formula(f"{col_letter}{r}")

                # Pricing sheet
                if "Pricing" in wb.sheetnames:
                    wb.remove(wb["Pricing"])
                pws = wb.create_sheet("Pricing")
                for r, row in enumerate(pricing_data, start=1):
                    for c, val in enumerate(row, start=1):
                        pws.cell(r, c, val)

                wb.save(dst)
                created.append(f)

            except Exception as e:
                st.warning(f"❌ Failed: {f} → {e}")

        st.success("🎉 New files created inside SOURCE folder!")
        st.dataframe(pd.DataFrame(created, columns=["New Files Created"]))






# =====================================================
# 6️⃣ COURIER COST UPDATER TOOL (FULLY MIRRORING WORKING SCRIPT)
# =====================================================
elif menu == "🚚 Courier Cost Updater":
    st.title("🚚 Courier Cost Updater Tool")
    
    testing_folder = st.text_input("📂 Testing Folder Path", r"D:\Sahil\Invoices\Python\Nov-2025\tODAY\07")
    cost_folder = st.text_input("📂 Courier Cost Folder Path", r"D:\Sahil\Invoices\Python\Courier Cost Final File ( From Ashwani )")
    
    source_files = live_file_status(testing_folder)
    start_btn = st.button("🚀 Start Updating Cost Data")

    if start_btn:
        progress = st.progress(0)
        logs = st.empty()
        processed, errors = [], []

        # --- Find cost file ---
        cost_file = next((f for f in os.listdir(cost_folder) if f.endswith(".xlsx") and not f.startswith("~$")), None)
        if not cost_file:
            st.error("❌ No cost file found in folder!")
        else:
            cost_file_path = os.path.join(cost_folder, cost_file)

            # --- Sheet1 ---
            df_cost1 = pd.read_excel(cost_file_path, sheet_name="Sheet1")
            df_cost1.columns.values[0] = 'AWB No'
            df_cost1['AWB No'] = df_cost1['AWB No'].astype(str)
            df_cost1_unique = df_cost1.iloc[:, :13].drop_duplicates(subset='AWB No', keep='first')
            cost_lookup1 = df_cost1_unique.set_index('AWB No').to_dict(orient='index')

            # --- Sheet2 ---
            df_cost2 = pd.read_excel(cost_file_path, sheet_name="Sheet2")
            df_cost2.columns.values[0] = 'AWB No'
            df_cost2['AWB No'] = df_cost2['AWB No'].astype(str)
            df_cost2_unique = df_cost2.iloc[:, :13].drop_duplicates(subset='AWB No', keep='first')
            cost_lookup2 = df_cost2_unique.set_index('AWB No').to_dict(orient='index')

            headers = [
                "Courier charged weight", "Courier Zone", "Freight", "RTO", "RTO Discount", "Reverse",
                "COD", "SDL", "Fuel", "QC", "Others", "Gross"
            ]
            col_start = 53  # BA column

            for idx, testing_file in enumerate(source_files, start=1):
                try:
                    testing_file_path = os.path.join(testing_folder, testing_file)
                    wb = openpyxl.load_workbook(testing_file_path)
                    ws = wb.active
                    max_row = ws.max_row

                    # 🔄 Clear BA-BS (old cost data)
                    for row in range(1, max_row + 1):
                        for col in range(col_start, col_start + len(headers)):
                            ws.cell(row=row, column=col).value = None

                    # 🧹 Blank entire row where column A is empty (from row 3 onward)
                    for row in range(3, max_row + 1):
                        if ws.cell(row=row, column=1).value in [None, ""]:
                            for col in range(1, col_start + len(headers)):
                                ws.cell(row=row, column=col).value = None

                    # 📝 Write headers in row 1 (BA-BM)
                    for i, header in enumerate(headers):
                        ws.cell(row=1, column=col_start + i).value = header

                    # 📝 Write cost data into rows 2+
                    for row in range(2, max_row + 1):
                        awb = ws.cell(row=row, column=4).value  # Column D
                        if awb and str(awb).strip() != "":
                            awb_str = str(awb).strip()
                            cost_row = None
                            source_sheet = None

                            # First check Sheet1, then Sheet2
                            if awb_str in cost_lookup1:
                                cost_row = list(cost_lookup1[awb_str].values())
                                source_sheet = "Sheet1"
                            elif awb_str in cost_lookup2:
                                cost_row = list(cost_lookup2[awb_str].values())
                                source_sheet = "Sheet2"

                            # Write cost data if found
                            if cost_row:
                                for col_offset, val in enumerate(cost_row):
                                    ws.cell(row=row, column=col_start + col_offset).value = val

                    # 💾 Save file
                    wb.save(testing_file_path)
                    processed.append(testing_file)
                    logs.write(f"✅ ({idx}/{len(source_files)}) Updated: {testing_file}")
                except Exception as e:
                    errors.append(f"{testing_file} → {e}")
                    logs.write(f"❌ ({idx}/{len(source_files)}) Error: {testing_file} → {e}")

                progress.progress(idx / len(source_files))

            st.success(f"🎉 Courier Cost Update Completed! Processed: {len(processed)} | Errors: {len(errors)}")
            if processed:
                st.dataframe(pd.DataFrame(processed, columns=["Processed"]))
            if errors:
                st.dataframe(pd.DataFrame(errors, columns=["Errors"]))

