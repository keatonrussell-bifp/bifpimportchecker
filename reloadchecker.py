import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO


# ==================================================
# Session State
# ==================================================
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None


# ==================================================
# Helpers
# ==================================================
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


# ==================================================
# PDF EXTRACTION (NUMERIC PROXIMITY – FINAL)
# ==================================================
def extract_lpn_and_pcs_from_pdfs(pdf_files):
    """
    PRODUCTION-GRADE parser.
    Does NOT rely on table layout, spacing, or labels.
    """

    lpns_set = set()
    pcs_map = {}

    for pdf in pdf_files:
        reader = PdfReader(pdf)

        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            # Extract ALL numeric tokens in order
            numbers = re.findall(r"\d+(?:\.\d+)?", text)

            cleaned = []
            for n in numbers:
                try:
                    cleaned.append(int(float(n)))
                except:
                    continue

            for i, val in enumerate(cleaned):
                # LPN = long numeric ID
                if 8 <= len(str(val)) <= 12:
                    lpn = str(val)
                    lpns_set.add(lpn)

                    # Look forward for PCS
                    pcs_candidate = None
                    for j in range(i + 1, min(i + 8, len(cleaned))):
                        pcs_val = cleaned[j]
                        if 1 <= pcs_val <= 5000:
                            pcs_candidate = pcs_val
                            break

                    if pcs_candidate is not None:
                        pcs_map[lpn] = str(pcs_candidate)

    return lpns_set, pcs_map


# ==================================================
# SKU LOGIC
# ==================================================
def map_description(grade):
    grade = str(grade).upper()
    if "APG" in grade:
        return "TAEDA PINE APG"
    if "DOG" in grade:
        return "DOG EAR"
    if re.search(r"\bIII/V\b|\bIII\b|\b3COM\b", grade):
        return "TAEDA PINE #3 COMMON"
    return "DOG EAR"


def sku_is_valid(val):
    if pd.isna(val):
        return False
    return str(val).strip() not in ("", "NAN", "NONE")


def load_sku_lookup(sku_file):
    REQUIRED = {
        "SKU": ["SKU"],
        "DESCRIPTION": ["DESCRIPTION", "DESC"],
        "THICKNESS": ["THICKNESS", "THK"],
        "WIDTH": ["WIDTH", "W"],
        "LENGTH": ["LENGTH", "LEN"]
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

    raise ValueError("SKU lookup missing required columns")


# ==================================================
# MAIN PROCESS
# ==================================================
def process_all(container_file, sku_file, pdf_files):

    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    # -------- PDF --------
    lpns_set, pcs_map = extract_lpn_and_pcs_from_pdfs(pdf_files)

    df["PDF LPN"] = df["PACKAGEID"].astype(str).apply(
        lambda x: x if x in lpns_set else ""
    )
    df["RECEIVE MATCH"] = df["PDF LPN"].apply(
        lambda x: "YES" if x else "NO"
    )

    df["PCS CHECK"] = df["PACKAGEID"].astype(str).apply(
        lambda x: pcs_map.get(x, "")
    )

    def pcs_match(container_pcs, pdf_pcs):
        try:
            return "YES" if int(container_pcs) == int(pdf_pcs) else "NO"
        except:
            return "NO"

    df["PCS MATCH"] = df.apply(
        lambda r: pcs_match(r.get("PCS", ""), r.get("PCS CHECK", "")),
        axis=1
    )

    # -------- SKU --------
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
    df["SKU"] = df["SKU"].apply(
        lambda x: str(x).strip() if sku_is_valid(x) else ""
    )

    # -------- ORDER --------
    audit_cols = [
        "PDF LPN", "RECEIVE MATCH",
        "PCS CHECK", "PCS MATCH",
        "SKU", "MATCH"
    ]
    existing = [c for c in audit_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing]
    return df[others + existing]


# ==================================================
# UI HIGHLIGHTING
# ==================================================
def highlight_mismatches(row):
    if (
        row["RECEIVE MATCH"] != "YES"
        or row["PCS MATCH"] != "YES"
        or row["MATCH"] != "YES"
    ):
        return ["background-color:#ffcccc"] * len(row)
    return [""] * len(row)


# ==================================================
# STREAMLIT UI
# ==================================================
st.set_page_config(page_title="BIFP Import Checker", layout="wide")
st.title("📦 BIFP SKU + Receive + PCS Audit Tool")

container_file = st.file_uploader("Upload Container List Excel", type="xlsx")
sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

if container_file and sku_file and pdf_files:
    if st.button("Run Full Process"):
        st.session_state.processed_df = process_all(
            container_file, sku_file, pdf_files
        )
        st.success("Processing complete")

        st.dataframe(
            st.session_state.processed_df.style.apply(
                highlight_mismatches, axis=1
            ),
            use_container_width=True
        )

        st.download_button(
            "⬇️ Download Audit Excel",
            to_excel_bytes(st.session_state.processed_df),
            container_file.name.replace(".xlsx", "_AUDIT_CHECK.xlsx")
        )

st.divider()

# ==================================================
# SALES ASSIST EXPORT
# ==================================================
st.header("Sales Assist Export")

sa_name = st.text_input(
    "Enter Sales Assist file name (no extension)",
    value="Sales_Assist"
)

if st.session_state.processed_df is not None and st.button("Generate Sales Assist Excel"):
    df = st.session_state.processed_df
    sa_df = pd.DataFrame({
        "SKU": df["SKU"],
        "Pieces": pd.to_numeric(df["PCS"], errors="coerce").fillna(0),
        "Quantity": pd.to_numeric(df["QTY"], errors="coerce").fillna(0),
        "QuantityUOM": "BF",
        "PriceUOM": "MBF",
        "PricePerUOM": 0,
        "OrderNumber": pd.to_numeric(
            df["ORDERNUMBER"].astype(str).str.split("-").str[0],
            errors="coerce"
        ).fillna(0),
        "ContainerNumber": df["CONTAINER"],
        "ReloadReference": "",
        "Identifier": pd.to_numeric(df["PACKAGEID"], errors="coerce").fillna(0),
        "ProFormaPrice": 0
    })

    st.download_button(
        "⬇️ Download Sales Assist Excel",
        to_excel_bytes(sa_df),
        f"{sa_name}.xlsx"
    )
