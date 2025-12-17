import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# --------------------------------------------------
# Utility: Excel download
# --------------------------------------------------
def to_excel_bytes(df):
    output = BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output


# --------------------------------------------------
# Stage 1: Receive Match Checker
# --------------------------------------------------
def extract_lpns_from_pdfs(pdf_files):
    lpns = set()
    pattern = re.compile(r"\b\d{8,12}\b")

    for pdf in pdf_files:
        reader = PdfReader(pdf)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue
            for m in pattern.findall(text):
                lpns.add(m)
    return lpns


def run_receive_match(excel_file, pdf_files):
    # Header is row 2
    df = pd.read_excel(excel_file, header=1, dtype=str).fillna("")
    df.columns = df.columns.astype(str).str.strip().str.upper()

    if "PACKAGEID" not in df.columns:
        raise ValueError("PACKAGEID column not found (expected on row 2)")

    lpns = extract_lpns_from_pdfs(pdf_files)

    pdf_lpn = []
    receive_match = []

    for pkg in df["PACKAGEID"]:
        pkg = str(pkg).strip()
        if pkg in lpns:
            pdf_lpn.append(pkg)
            receive_match.append("YES")
        else:
            pdf_lpn.append("")
            receive_match.append("NO")

    df["PDF LPN"] = pdf_lpn
    df["RECEIVE MATCH"] = receive_match

    return df


# --------------------------------------------------
# Stage 2: SKU Adder
# --------------------------------------------------
def map_description(grade: str) -> str:
    grade = str(grade).upper().strip()

    if "APG" in grade:
        return "TAEDA PINE APG"
    if "DOG" in grade:
        return "DOG EAR"
    if re.search(r"\bIII/V\b|\bIII\b", grade):
        return "TAEDA PINE #3 COMMON"

    # Default
    return "DOG EAR"


def normalize_container_headers(df):
    header_row = None
    for i in range(min(20, len(df))):
        row = df.iloc[i].astype(str).str.upper()
        if "GRADE" in row.values:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate header row containing GRADE")

    df.columns = df.iloc[header_row].astype(str).str.strip().str.upper()
    return df.iloc[header_row + 1:].reset_index(drop=True)


def run_sku_adder(container_df, sku_lookup_file):
    # Load SKU lookup
    sku_xls = pd.ExcelFile(sku_lookup_file)
    sku_df = None

    for sheet in sku_xls.sheet_names:
        temp = sku_xls.parse(sheet, dtype=str)
        temp.columns = temp.columns.astype(str).str.strip().str.upper()
        if {"SKU", "DESCRIPTION", "THICKNESS", "WIDTH", "LENGTH"}.issubset(temp.columns):
            sku_df = temp
            break

    if sku_df is None:
        raise ValueError("SKU lookup must contain SKU, DESCRIPTION, THICKNESS, WIDTH, LENGTH")

    sku_df = sku_df.fillna("")
    sku_df["DESCRIPTION"] = sku_df["DESCRIPTION"].str.upper().str.strip()

    sku_df["MATCH KEY"] = (
        sku_df["DESCRIPTION"] + "|" +
        sku_df["THICKNESS"] + "|" +
        sku_df["WIDTH"] + "|" +
        sku_df["LENGTH"]
    )

    # Normalize container headers again (works for receive-match output too)
    raw_df = container_df.copy()
    raw_df.columns = raw_df.columns.astype(str)
    cont_df = normalize_container_headers(raw_df).fillna("")

    cont_df["MAPPED DESCRIPTION"] = cont_df["GRADE"].apply(map_description)

    cont_df["MATCH KEY"] = (
        cont_df["MAPPED DESCRIPTION"] + "|" +
        cont_df["THICKNESS"] + "|" +
        cont_df["WIDTH"] + "|" +
        cont_df["LENGTH"]
    )

    final_df = cont_df.merge(
        sku_df[["SKU", "MATCH KEY"]],
        how="left",
        on="MATCH KEY"
    )

    return final_df


# --------------------------------------------------
# Streamlit UI
# --------------------------------------------------
st.set_page_config(page_title="Receive Match + SKU Adder", layout="wide")
st.title("📦 Receive Match Checker + SKU Adder")

st.markdown(
    """
This app supports a **two-stage workflow**:

1. **Receive Match Checker**  
   Match `PACKAGEID` from Excel against LPNs in PDFs.

2. **SKU Adder (Optional)**  
   Add SKUs using the output from Receive Match Checker.
"""
)

st.divider()

# ======================
# Stage 1 UI
# ======================
st.header("Step 1️⃣ Receive Match Checker")

excel_rm = st.file_uploader(
    "Upload Container Excel (PACKAGEID on row 2)",
    type=["xlsx"],
    key="rm_excel"
)

pdfs_rm = st.file_uploader(
    "Upload PDF files",
    type=["pdf"],
    accept_multiple_files=True,
    key="rm_pdfs"
)

rm_df = None

if excel_rm and pdfs_rm:
    if st.button("Run Receive Match"):
        try:
            rm_df = run_receive_match(excel_rm, pdfs_rm)
            st.success("Receive Match completed")
            st.dataframe(rm_df.head(50), use_container_width=True)

            st.download_button(
                "⬇️ Download Receive Match Excel",
                data=to_excel_bytes(rm_df),
                file_name=excel_rm.name.replace(".xlsx", "_RECEIVE_MATCH.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(str(e))

st.divider()

# ======================
# Stage 2 UI
# ======================
st.header("Step 2️⃣ SKU Adder (Optional)")

sku_lookup = st.file_uploader(
    "Upload SKU Lookup Excel",
    type=["xlsx"],
    key="sku_lookup"
)

sku_input_excel = st.file_uploader(
    "Upload Receive Match Excel (or original container list)",
    type=["xlsx"],
    key="sku_input"
)

if sku_lookup and sku_input_excel:
    if st.button("Run SKU Adder"):
        try:
            # Read the uploaded SKU input as raw DataFrame
            raw_df = pd.read_excel(sku_input_excel, header=None, dtype=str)
            final_sku_df = run_sku_adder(raw_df, sku_lookup)

            st.success("SKU Adder completed")
            st.dataframe(final_sku_df.head(50), use_container_width=True)

            st.download_button(
                "⬇️ Download SKU Added Excel",
                data=to_excel_bytes(final_sku_df),
                file_name=sku_input_excel.name.replace(".xlsx", "_SKU_ADDED.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(str(e))
