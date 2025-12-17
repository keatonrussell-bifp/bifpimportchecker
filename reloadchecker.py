import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# --------------------------------------------------
# Helpers
# --------------------------------------------------
def to_excel_bytes(df):
    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    return bio


def normalize_headers(df):
    header_row = None
    for i in range(min(20, len(df))):
        if "GRADE" in df.iloc[i].astype(str).str.upper().values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not locate header row containing GRADE")
    df.columns = df.iloc[header_row].astype(str).str.strip().str.upper()
    return df.iloc[header_row + 1:].reset_index(drop=True)


# --------------------------------------------------
# Receive Match
# --------------------------------------------------
def extract_lpns_from_pdfs(pdf_files):
    lpns = set()
    pattern = re.compile(r"\b\d{8,12}\b")
    for pdf in pdf_files:
        reader = PdfReader(pdf)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                lpns.update(pattern.findall(text))
    return lpns


def run_receive_match(excel_file, pdf_files):
    df = pd.read_excel(excel_file, header=1, dtype=str).fillna("")
    df.columns = df.columns.str.upper().str.strip()

    if "PACKAGEID" not in df.columns:
        raise ValueError("PACKAGEID column not found (row 2 expected)")

    lpns = extract_lpns_from_pdfs(pdf_files)

    df["PDF LPN"] = df["PACKAGEID"].apply(lambda x: x if x in lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].apply(lambda x: "YES" if x in lpns else "NO")
    return df


# --------------------------------------------------
# SKU Adder
# --------------------------------------------------
def map_description(grade):
    grade = str(grade).upper()
    if "APG" in grade:
        return "TAEDA PINE APG"
    if "DOG" in grade:
        return "DOG EAR"
    if re.search(r"\bIII/V\b|\bIII\b", grade):
        return "TAEDA PINE #3 COMMON"
    return "DOG EAR"


def run_sku_adder(container_raw_df, sku_file):
    container_df = normalize_headers(container_raw_df).fillna("")

    # Load SKU lookup
    xls = pd.ExcelFile(sku_file)
    sku_df = None
    for sheet in xls.sheet_names:
        tmp = xls.parse(sheet, dtype=str)
        tmp.columns = tmp.columns.str.upper().str.strip()
        if {"SKU", "DESCRIPTION", "THICKNESS", "WIDTH", "LENGTH"}.issubset(tmp.columns):
            sku_df = tmp
            break

    if sku_df is None:
        raise ValueError("SKU lookup missing required columns")

    sku_df = sku_df.fillna("")
    sku_df["DESCRIPTION"] = sku_df["DESCRIPTION"].str.upper().str.strip()

    sku_df["MATCH KEY"] = (
        sku_df["DESCRIPTION"] + "|" +
        sku_df["THICKNESS"] + "|" +
        sku_df["WIDTH"] + "|" +
        sku_df["LENGTH"]
    )

    container_df["MAPPED DESCRIPTION"] = container_df["GRADE"].apply(map_description)
    container_df["MATCH KEY"] = (
        container_df["MAPPED DESCRIPTION"] + "|" +
        container_df["THICKNESS"] + "|" +
        container_df["WIDTH"] + "|" +
        container_df["LENGTH"]
    )

    out = container_df.merge(
        sku_df[["SKU", "MATCH KEY"]],
        how="left",
        on="MATCH KEY"
    )

    out["MATCH"] = out["SKU"].apply(lambda x: "YES" if str(x).strip() else "NO")
    return out


# --------------------------------------------------
# Sales Assist Generator
# --------------------------------------------------
def generate_sales_assist(df):
    order_number = (
        df.get("ORDERNUMBER", "")
        .astype(str)
        .str.split("-")
        .str[0]
    )

    return pd.DataFrame({
        "SKU": df.get("SKU", ""),
        "Pieces": pd.to_numeric(df.get("PCS", 0), errors="coerce").fillna(0),
        "Quantity": pd.to_numeric(df.get("QTY", 0), errors="coerce").fillna(0),
        "QuantityUOM": "BF",
        "PriceUOM": "MBF",
        "PricePerUOM": 0,
        "OrderNumber": pd.to_numeric(order_number, errors="coerce").fillna(0),
        "ContainerNumber": df.get("CONTAINER", ""),
        "ReloadReference": "",
        "Identifier": pd.to_numeric(df.get("PACKAGEID", 0), errors="coerce").fillna(0),
        "ProFormaPrice": 0
    })


# --------------------------------------------------
# Streamlit UI
# --------------------------------------------------
st.set_page_config(page_title="Receive Match + SKU + Sales Assist", layout="wide")
st.title("📦 Receive Match → SKU Adder → Sales Assist")

# ------------------ Step 1 ------------------
st.header("Step 1️⃣ Receive Match Checker")

rm_excel = st.file_uploader("Upload Container Excel (PACKAGEID on row 2)", type="xlsx", key="rm_excel")
rm_pdfs = st.file_uploader("Upload PDFs", type="pdf", accept_multiple_files=True, key="rm_pdfs")

rm_df = None
if rm_excel and rm_pdfs and st.button("Run Receive Match"):
    rm_df = run_receive_match(rm_excel, rm_pdfs)
    st.success("Receive Match completed")
    st.dataframe(rm_df.head(50), use_container_width=True)
    st.download_button(
        "⬇️ Download Receive Match Excel",
        to_excel_bytes(rm_df),
        rm_excel.name.replace(".xlsx", "_RECEIVE_MATCH.xlsx")
    )

st.divider()

# ------------------ Step 2 ------------------
st.header("Step 2️⃣ SKU Adder")

sku_lookup = st.file_uploader("Upload SKU Lookup Excel", type="xlsx", key="sku_lookup")
sku_input = st.file_uploader("Upload Receive Match Excel (or original)", type="xlsx", key="sku_input")

sku_df = None
if sku_lookup and sku_input and st.button("Run SKU Adder"):
    raw_df = pd.read_excel(sku_input, header=None, dtype=str)
    sku_df = run_sku_adder(raw_df, sku_lookup)
    st.success("SKU Adder completed")
    st.dataframe(sku_df.head(50), use_container_width=True)
    st.download_button(
        "⬇️ Download SKU Added Excel",
        to_excel_bytes(sku_df),
        sku_input.name.replace(".xlsx", "_SKU_ADDED.xlsx")
    )

st.divider()

# ------------------ Step 3 ------------------
st.header("Step 3️⃣ Sales Assist Report")

if sku_df is not None and st.button("Generate Sales Assist Excel"):
    sa_df = generate_sales_assist(sku_df)
    st.success("Sales Assist report generated")
    st.dataframe(sa_df.head(50), use_container_width=True)
    st.download_button(
        "⬇️ Download Sales Assist Excel",
        to_excel_bytes(sa_df),
        "Sales_Assist.xlsx"
    )
