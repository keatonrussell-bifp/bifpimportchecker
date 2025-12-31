import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# --------------------------------------------------
# Session State Init
# --------------------------------------------------
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None


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


def norm_id(x):
    """
    Normalize IDs like:
    - 7052583281.0 -> 7052583281
    - '7052583281  -> 7052583281
    - 7.05258E+09  -> 7052583281
    """
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.startswith("'"):
        s = s[1:].strip()
    s = s.replace(",", "")

    # handle scientific/float-like safely
    try:
        f = float(s)
        if f.is_integer():
            s = str(int(f))
    except:
        pass

    m = re.search(r"\d{8,12}", s)
    return m.group(0) if m else s


def norm_int_str(x):
    """Normalize integer-like strings; blank if not parseable."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.startswith("'"):
        s = s[1:].strip()
    s = s.replace(",", "")
    try:
        return str(int(float(s)))
    except:
        return ""


# --------------------------------------------------
# LPN / PDF Logic (KEEP YOUR WORKING LPN EXTRACTOR)
# --------------------------------------------------
def extract_lpns_from_pdfs(pdf_files):
    lpns = set()
    pattern = re.compile(r"\b\d{8,12}\b")

    for pdf in pdf_files:
        # Use BytesIO(pdf.getvalue()) so we never get file-pointer issues
        reader = PdfReader(BytesIO(pdf.getvalue()))
        for page in reader.pages:
            text = page.extract_text()
            if text:
                lpns.update(pattern.findall(text))
    return lpns


# --------------------------------------------------
# PCS / PIECES from PDF Logic (NEW, BASED ON YOUR EXTRACTED TEXT)
# --------------------------------------------------
def extract_pcs_map_from_pdfs(pdf_files):
    """
    Extract PCS/PIECES by matching the pattern you pasted:
      LPN  PIECES  TOTAL_LBS
    Example:
      7052583281 240 1120.0000

    Returns dict: {LPN: PIECES}
    """
    pcs_map = {}

    # LPN (8–12 digits) + PIECES (1–5 digits) + TOTAL LBS (decimal)
    row_pattern = re.compile(r"\b(\d{8,12})\b\s+(\d{1,5})\b\s+\d+(?:\.\d+)?")

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            # Collapse whitespace so row regex works even if the PDF breaks lines oddly
            flat = re.sub(r"\s+", " ", text)

            for lpn, pieces in row_pattern.findall(flat):
                # keep first seen; if repeats exist they should be same
                if lpn not in pcs_map:
                    pcs_map[lpn] = pieces

    return pcs_map


# --------------------------------------------------
# SKU Logic
# --------------------------------------------------
def map_description(grade):
    grade = str(grade).upper()
    if "APG" in grade:
        return "TAEDA PINE APG"
    if "DOG" in grade:
        return "DOG EAR"
    if re.search(r"\bIII/V\b|\bIII\b|\b3COM\b", grade):
        return "TAEDA PINE #3 COMMON"
    return "DOG EAR"


def load_sku_lookup(sku_file):
    # Slightly more tolerant than strict, but still safe
    REQUIRED = {
        "SKU": ["SKU"],
        "DESCRIPTION": ["DESCRIPTION", "DESC", "PRODUCT DESCRIPTION", "GRADE DESC"],
        "THICKNESS": ["THICKNESS", "THK"],
        "WIDTH": ["WIDTH", "W"],
        "LENGTH": ["LENGTH", "LEN", "L"]
    }

    xls = pd.ExcelFile(sku_file)
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, dtype=str)
        df.columns = df.columns.str.upper().str.strip()

        col_map = {}
        for canon, aliases in REQUIRED.items():
            for a in aliases:
                if a in df.columns:
                    col_map[a] = canon
                    break

        if set(col_map.values()) == set(REQUIRED.keys()):
            df = df.rename(columns=col_map).fillna("")
            df["DESCRIPTION"] = df["DESCRIPTION"].str.upper().str.strip()
            df["MATCH KEY"] = (
                df["DESCRIPTION"] + "|" +
                df["THICKNESS"] + "|" +
                df["WIDTH"] + "|" +
                df["LENGTH"]
            )
            return df

    raise ValueError("SKU lookup missing required columns (SKU/DESCRIPTION/THICKNESS/WIDTH/LENGTH).")


def sku_is_valid(val):
    if pd.isna(val):
        return False
    val = str(val).strip().upper()
    return val not in ("", "NAN", "NONE")


# --------------------------------------------------
# Combined Processor
# --------------------------------------------------
def process_all(container_file, sku_file, pdf_files):
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    # Validate required columns in container list
    required_cols = {"PACKAGEID", "PCS", "GRADE", "THICKNESS", "WIDTH", "LENGTH"}
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Container list missing required columns: {missing}")

    # Normalize keys
    df["PACKAGEID"] = df["PACKAGEID"].apply(norm_id)
    df["PCS"] = df["PCS"].apply(norm_int_str)

    # --- Receive Match (LPN)
    lpns = extract_lpns_from_pdfs(pdf_files)
    df["PDF LPN"] = df["PACKAGEID"].apply(lambda x: x if x in lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].apply(lambda x: "YES" if x in lpns else "NO")

    # --- PCS Check
    pcs_map = extract_pcs_map_from_pdfs(pdf_files)
    df["PCS CHECK"] = df["PACKAGEID"].apply(lambda x: pcs_map.get(x, ""))

    def pcs_match(container_pcs, pdf_pcs):
        try:
            return "YES" if int(container_pcs) == int(pdf_pcs) else "NO"
        except:
            return "NO"

    df["PCS MATCH"] = df.apply(
        lambda r: pcs_match(r.get("PCS", ""), r.get("PCS CHECK", "")),
        axis=1
    )

    # --- SKU Match
    sku_df = load_sku_lookup(sku_file)
    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)
    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|" +
        df["THICKNESS"] + "|" +
        df["WIDTH"] + "|" +
        df["LENGTH"]
    )

    df = df.merge(
        sku_df[["SKU", "MATCH KEY"]],
        how="left",
        on="MATCH KEY"
    )

    df["MATCH"] = df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")

    # --- Audit-friendly ordering (put audit columns at the end)
    audit_cols = ["PDF LPN", "RECEIVE MATCH", "PCS CHECK", "PCS MATCH", "SKU", "MATCH"]
    existing_audit = [c for c in audit_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_audit]
    df = df[others + existing_audit]

    return df


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
# UI Styling: Red rows if any mismatch
# --------------------------------------------------
def highlight_mismatches(row):
    # Any NO triggers red row
    if (
        row.get("RECEIVE MATCH") != "YES"
        or row.get("PCS MATCH") != "YES"
        or row.get("MATCH") != "YES"
    ):
        return ["background-color: #ffcccc"] * len(row)
    return [""] * len(row)


# --------------------------------------------------
# Streamlit UI
# --------------------------------------------------
st.set_page_config(page_title="SKU + Receive + PCS Match + Sales Assist", layout="wide")
st.title("📦 SKU + Receive + PCS Match + Sales Assist Generator")

container_file = st.file_uploader("Upload Container List Excel", type="xlsx")
sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

if container_file and sku_file and pdf_files:
    if st.button("Run Full Process"):
        st.session_state.processed_df = process_all(
            container_file, sku_file, pdf_files
        )
        st.success("Full process completed")

        # Styled view (red mismatches)
        styled = st.session_state.processed_df.style.apply(highlight_mismatches, axis=1)
        st.dataframe(styled, use_container_width=True)

        st.download_button(
            "⬇️ Download SKU + Receive + PCS Match Excel",
            to_excel_bytes(st.session_state.processed_df),
            container_file.name.replace(".xlsx", "_SKU_RECEIVE_PCS_MATCH.xlsx")
        )

st.divider()

# ---------------- Sales Assist ----------------
st.header("Sales Assist Export")

sa_name = st.text_input(
    "Enter Sales Assist file name (no extension)",
    value="Sales_Assist"
)

if st.session_state.processed_df is None:
    st.info("ℹ️ Run the **Full Process** above before generating Sales Assist.")

if st.session_state.processed_df is not None and st.button("Generate Sales Assist Excel"):
    sa_df = generate_sales_assist(st.session_state.processed_df)
    st.success("Sales Assist report generated")

    st.download_button(
        "⬇️ Download Sales Assist Excel",
        to_excel_bytes(sa_df),
        f"{sa_name}.xlsx"
    )
