import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# --------------------------------------------------
# Session State
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


# --------------------------------------------------
# PDF Extraction
# --------------------------------------------------
def extract_lpn_and_pcs_from_pdfs(pdf_files):
    """
    Returns:
      lpns_set -> set of PACKAGEIDs found
      pcs_map  -> {PACKAGEID: PCS}
    """
    lpns = set()
    pcs_map = {}

    lpn_pattern = re.compile(r"\b\d{8,12}\b")
    pcs_pattern = re.compile(r"\bPCS[:\s]*([0-9]+)\b", re.IGNORECASE)

    for pdf in pdf_files:
        reader = PdfReader(pdf)
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            found_lpns = lpn_pattern.findall(text)
            pcs_match = pcs_pattern.search(text)

            if found_lpns:
                lpns.update(found_lpns)

                if pcs_match:
                    pcs_val = pcs_match.group(1)
                    for lpn in found_lpns:
                        pcs_map[lpn] = pcs_val

    return lpns, pcs_map


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
    xls = pd.ExcelFile(sku_file)
    for sheet in xls.sheet_names:
        df = xls.parse(sheet, dtype=str)
        df.columns = df.columns.str.upper().str.strip()
        if {"SKU", "DESCRIPTION", "THICKNESS", "WIDTH", "LENGTH"}.issubset(df.columns):
            df = df.fillna("")
            df["DESCRIPTION"] = df["DESCRIPTION"].str.upper().str.strip()
            df["MATCH KEY"] = (
                df["DESCRIPTION"] + "|" +
                df["THICKNESS"] + "|" +
                df["WIDTH"] + "|" +
                df["LENGTH"]
            )
            return df
    raise ValueError("SKU lookup missing required columns")


def sku_is_valid(val):
    if pd.isna(val):
        return False
    val = str(val).strip().upper()
    return val not in ("", "NAN", "NONE")


# --------------------------------------------------
# Main Processor
# --------------------------------------------------
def process_all(container_file, sku_file, pdf_files):

    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    # -------- PDF READ --------
    lpns_set, pcs_map = extract_lpn_and_pcs_from_pdfs(pdf_files)

    # -------- RECEIVE CHECK --------
    df["PDF LPN"] = df["PACKAGEID"].apply(lambda x: x if str(x) in lpns_set else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].apply(
        lambda x: "YES" if str(x) in lpns_set else "NO"
    )

    # -------- PCS CHECK --------
    df["PDF PCS"] = df["PACKAGEID"].apply(
        lambda x: pcs_map.get(str(x), "")
    )

    def pcs_match(container_pcs, pdf_pcs):
        try:
            return "YES" if int(container_pcs) == int(pdf_pcs) else "NO"
        except:
            return "NO"

    df["PCS MATCH"] = df.apply(
        lambda r: pcs_match(r.get("PCS", ""), r.get("PDF PCS", "")),
        axis=1
    )

    # -------- SKU MATCH --------
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

    df["MATCH"] = df["SKU"].apply(
        lambda x: "YES" if sku_is_valid(x) else "NO"
    )

    # Optional cleanup: blank invalid SKUs
    df["SKU"] = df["SKU"].apply(
        lambda x: str(x).strip() if sku_is_valid(x) else ""
    )

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
# Streamlit UI
# --------------------------------------------------
st.set_page_config(page_title="SKU + Receive + PCS Checker", layout="wide")
st.title("📦 SKU + Receive + PCS Validation Tool")

container_file = st.file_uploader("Upload Container List Excel", type="xlsx")
sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

if container_file and sku_file and pdf_files:
    if st.button("Run Full Process"):
        st.session_state.processed_df = process_all(
            container_file, sku_file, pdf_files
        )
        st.success("Processing completed")
        st.dataframe(
            st.session_state.processed_df.head(50),
            use_container_width=True
        )

        st.download_button(
            "⬇️ Download SKU + Receive + PCS Check Excel",
            to_excel_bytes(st.session_state.processed_df),
            container_file.name.replace(".xlsx", "_AUDIT_CHECK.xlsx")
        )

st.divider()

# ---------------- Sales Assist ----------------
st.header("Sales Assist Export")

sa_name = st.text_input(
    "Enter Sales Assist file name (no extension)",
    value="Sales_Assist"
)

if st.session_state.processed_df is None:
    st.info("ℹ️ Run the full process before generating Sales Assist.")

if st.session_state.processed_df is not None and st.button("Generate Sales Assist Excel"):
    sa_df = generate_sales_assist(st.session_state.processed_df)
    st.success("Sales Assist report generated")

    st.download_button(
        "⬇️ Download Sales Assist Excel",
        to_excel_bytes(sa_df),
        f"{sa_name}.xlsx"
    )
