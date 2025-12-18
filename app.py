import streamlit as st
import pandas as pd
import io
import zipfile
import requests
from math import radians, sin, cos, sqrt, atan2
from concurrent.futures import ThreadPoolExecutor
import os

st.set_page_config(page_title="📦 Helper Portal - Sahil", layout="wide")
st.title("📦 Helper Portal - Sahil")

# =====================================================
# UTILITIES
# =====================================================
METRO_RANGES = [
    range(110001, 110099), range(400001, 400105), range(700001, 700105),
    range(600001, 600119), range(560001, 560108), range(500001, 500099),
    range(380001, 380062), range(411001, 411063), range(122001, 122019)
]

def is_metro(pin):
    return any(int(pin) in r for r in METRO_RANGES)

def get_location(pin):
    try:
        r = requests.get(
            f"https://api.postalpincode.in/pincode/{pin}",
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=10
        )
        d = r.json()
        if d[0]["Status"].lower() == "success":
            po = d[0]["PostOffice"][0]
            return po.get("Name",""), po.get("District",""), po.get("State","")
    except:
        pass
    return "N/A", "N/A", "N/A"

def get_latlon(pin):
    try:
        r = requests.get(
            f"https://nominatim.openstreetmap.org/search?postalcode={pin}&country=India&format=json",
            headers={"User-Agent": "Mozilla/5.0"}
        )
        d = r.json()
        if d:
            return float(d[0]["lat"]), float(d[0]["lon"])
    except:
        pass
    return None, None

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1))*cos(radians(lat2))*sin(dlon/2)**2
    return round(R * 2 * atan2(sqrt(a), sqrt(1 - a)), 2)

def classify_zone(fpin, tpin, fd, fs, td, ts):
    special = {
        "himachal pradesh","karnataka","jammu & kashmir","west bengal",
        "assam","manipur","mizoram","nagaland","tripura",
        "meghalaya","sikkim","arunachal pradesh"
    }
    if fpin == tpin:
        return "LOCAL"
    if fd.lower() == td.lower() and fd != "N/A":
        return "LOCAL"
    if is_metro(fpin) and is_metro(tpin):
        return "METRO"
    if fs.lower() == ts.lower():
        return "REGIONAL"
    if fs.lower() in special or ts.lower() in special:
        return "SPECIAL"
    return "ROI"

def process(row):
    f = str(row["from_pincode"])
    t = str(row["to_pincode"])
    fc, fd, fs = get_location(f)
    tc, td, ts = get_location(t)
    lat1, lon1 = get_latlon(f)
    lat2, lon2 = get_latlon(t)
    dist = haversine(lat1, lon1, lat2, lon2) if None not in [lat1, lon1, lat2, lon2] else "N/A"
    zone = classify_zone(f, t, fd, fs, td, ts)

    return {
        "From Pincode": f,
        "To Pincode": t,
        "From City": fc,
        "From State": fs,
        "To City": tc,
        "To State": ts,
        "Distance (KM)": dist,
        "Zone": zone
    }

def extract_files(uploaded_files, zip_file, extensions):
    files = []
    if uploaded_files:
        files.extend(uploaded_files)

    if zip_file:
        z = zipfile.ZipFile(zip_file)
        for n in z.namelist():
            if n.lower().endswith(extensions):
                buf = io.BytesIO(z.read(n))
                buf.name = os.path.basename(n)
                files.append(buf)
    return files

def make_zip(file_buffers):
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zf:
        for name, buf in file_buffers:
            zf.writestr(name, buf.getvalue())
    zip_buffer.seek(0)
    return zip_buffer

def read_any_file(f):
    name = f.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(f)
    elif name.endswith(".xlsx"):
        return pd.read_excel(f)
    elif name.endswith(".xlsb"):
        return pd.read_excel(f, engine="pyxlsb")
    else:
        return None

# =====================================================
# SIDEBAR
# =====================================================
tool = st.sidebar.selectbox("Choose Tool", [
    "Data Compiler",
    "Files Splitter",
    "Pincode Zone + Distance",
    "Data Cleaner & Summary",
    "Create Folders from List",
    "CSV → XLSX Converter",
    "Merchant Auto Rename"
])

# =====================================================
# DATA COMPILER (CSV + XLSX + XLSB)
# =====================================================
if tool == "Data Compiler":
    st.header("📁 Data Compiler (ALL Formats)")
    files = st.file_uploader(
        "Upload CSV / XLSX / XLSB Files",
        type=["csv", "xlsx", "xlsb"],
        accept_multiple_files=True
    )

    if files:
        dfs = []
        for f in files:
            df = read_any_file(f)
            if df is not None:
                df["Source File"] = f.name
                dfs.append(df)

        if dfs:
            final = pd.concat(dfs, ignore_index=True).fillna("")
            st.dataframe(final, use_container_width=True)
            st.download_button(
                "⬇️ Download Compiled CSV",
                final.to_csv(index=False).encode(),
                "compiled.csv"
            )

# =====================================================
# FILES SPLITTER
# =====================================================
elif tool == "Files Splitter":
    st.header("📂 Files Splitter")
    f = st.file_uploader("Upload CSV / Excel", type=["csv","xlsx","xlsb"])
    if f:
        df = read_any_file(f)
        st.dataframe(df.head(), use_container_width=True)
        col = st.selectbox("Split By Column", df.columns)

        if st.button("Split & Download"):
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w") as z:
                for v in df[col].dropna().unique():
                    d = df[df[col] == v]
                    z.writestr(f"{v}.csv", d.to_csv(index=False))
            st.download_button("⬇️ Download ZIP", zip_buf.getvalue(), "split.zip")

# =====================================================
# PINCODE ZONE + DISTANCE
# =====================================================
elif tool == "Pincode Zone + Distance":
    st.header("📍 Pincode Zone + Distance")
    mode = st.radio("Mode", ["Upload File", "Manual Entry"])

    if mode == "Upload File":
        f = st.file_uploader("Upload CSV/XLSX", type=["csv","xlsx"])
        if f:
            df = read_any_file(f)
            if {"from_pincode","to_pincode"}.issubset(df.columns):
                with ThreadPoolExecutor(max_workers=10) as ex:
                    out = list(ex.map(process, df.to_dict("records")))
                res = pd.DataFrame(out)
                st.dataframe(res, use_container_width=True)
                st.download_button("⬇️ Download CSV", res.to_csv(index=False).encode(), "zones.csv")

    else:
        txt = st.text_area("from,to per line")
        if txt:
            pairs = [l.split(",") for l in txt.splitlines() if "," in l]
            df = pd.DataFrame(pairs, columns=["from_pincode","to_pincode"])
            out = [process(r) for r in df.to_dict("records")]
            res = pd.DataFrame(out)
            st.dataframe(res, use_container_width=True)

# =====================================================
# DATA CLEANER
# =====================================================
elif tool == "Data Cleaner & Summary":
    st.header("🧹 Data Cleaner")
    f = st.file_uploader("Upload File", type=["csv","xlsx","xlsb"])
    if f:
        df = read_any_file(f)
        df = df.drop_duplicates().dropna(how="all")
        df.columns = [c.strip().title() for c in df.columns]
        st.dataframe(df, use_container_width=True)
        st.download_button("⬇️ Download Cleaned CSV", df.to_csv(index=False).encode(), "cleaned.csv")

# =====================================================
# CREATE FOLDERS
# =====================================================
elif tool == "Create Folders from List":
    st.header("📂 Create Folders")
    txt = st.text_area("Folder names (one per line)")
    if st.button("Create ZIP"):
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w") as z:
            for n in txt.splitlines():
                if n.strip():
                    z.writestr(f"{n.strip()}/", "")
        st.download_button("⬇️ Download ZIP", zip_buf.getvalue(), "folders.zip")

# =====================================================
# CSV → XLSX
# =====================================================
elif tool == "CSV → XLSX Converter":
    st.header("🔄 CSV → XLSX")
    files = st.file_uploader("Upload CSV", type="csv", accept_multiple_files=True)
    zip_f = st.file_uploader("OR ZIP", type="zip")

    if st.button("Convert"):
        csvs = extract_files(files, zip_f, (".csv",))
        output = []
        for f in csvs:
            df = pd.read_csv(f)
            buf = io.BytesIO()
            df.to_excel(buf, index=False)
            buf.seek(0)
            output.append((f.name.replace(".csv",".xlsx"), buf))
        st.download_button("⬇️ Download ZIP", make_zip(output), "converted.xlsx.zip")

# =====================================================
# MERCHANT AUTO RENAME
# =====================================================
elif tool == "Merchant Auto Rename":
    st.header("📝 Merchant Auto Rename")
    files = st.file_uploader("Upload XLSX", type="xlsx", accept_multiple_files=True)
    zip_f = st.file_uploader("OR ZIP", type="zip")

    if st.button("Rename"):
        excels = extract_files(files, zip_f, (".xlsx",))
        output = []
        for f in excels:
            df = pd.read_excel(f, header=None)
            name = f.name
            if str(df.iloc[0,0]).strip().lower() == "client name":
                name = str(df.iloc[1,0]).strip() + ".xlsx"

            buf = io.BytesIO()
            df.to_excel(buf, index=False, header=False)
            buf.seek(0)
            output.append((name, buf))

        st.download_button("⬇️ Download ZIP", make_zip(output), "renamed.zip")
