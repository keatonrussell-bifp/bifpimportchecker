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


def norm_digits(x):
    """Normalize things like 7052583281.0 -> 7052583281 and strip whitespace."""
    s = "" if pd.isna(x) else str(x).strip()
    s = re.sub(r"\.0$", "", s)
    return s


# --------------------------------------------------
# PDF Logic (ONE PASS: LPN + PIECES)
# --------------------------------------------------
def extract_lpns_and_pieces_from_pdfs(pdf_files):
    """
    Returns:
      lpns_set: set of LPNs found (8-12 digits)
      pieces_map: dict {LPN: PIECES} extracted by proximity after each LPN
    NOTE: We read each PDF from BytesIO(pdf.getvalue()) so we never hit EOF pointer issues.
    """
    lpns = set()
    pieces_map = {}

    lpn_pat = re.compile(r"\b\d{8,12}\b")
    num_token_pat = re.compile(r"\d+(?:\.\d+)?")  # integers + decimals

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))

        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            # --- LPNs (this is the SAME logic you had before, just done in the shared loop)
            page_lpns = lpn_pat.findall(text)
            lpns.update(page_lpns)

            # --- PIECES per LPN (works even when table spacing/lines are weird)
            # For every LPN occurrence, look ahead for the first *integer* token (no decimal) that looks like a PCS count.
            for m in lpn_pat.finditer(text):
                lpn = m.group(0)

                # Look ahead a bit after the LPN occurrence
                tail = text[m.end(): m.end() + 250]
                tokens = num_token_pat.findall(tail)

                pieces_val = None
                for t in tokens:
                    # Skip decimals like 1120.0000 (TOTAL LBS)
                    if "." in t:
                        continue
                    try:
                        v = int(t)
                    except:
                        continue

                    # Reasonable bounds for PIECES
                    if 1 <= v <= 5000:
                        pieces_val = str(v)
                        break

                if pieces_val and lpn not in pieces_map:
                    pieces_map[lpn] = pieces_val

    return lpns, pieces_map


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
# Combined Processor
# --------------------------------------------------
def process_all(container_file, sku_file, pdf_files):
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    # Normalize key fields
    if "PACKAGEID" in df.columns:
        df["PACKAGEID"] = df["PACKAGEID"].apply(norm_digits)
    if "PCS" in df.columns:
        df["PCS"] = df["PCS"].apply(norm_digits)

    # --- Receive Match + PCS Check (from PDFs in one pass)
    lpns, pieces_map = extract_lpns_and_pieces_from_pdfs(pdf_files)

    df["PDF LPN"] = df["PACKAGEID"].apply(lambda x: x if x in lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].apply(lambda x: "YES" if x in lpns else "NO")

    # PCS CHECK = value from PDFs
    df["PCS CHECK"] = df["PACKAGEID"].apply(lambda x: pieces_map.get(x, ""))

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

    # Optional: audit-friendly ordering of the key columns at the end
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

        # quick stats
        dfp = st.session_state.processed_df
        st.write(
            {
                "Rows": len(dfp),
                "Receive NO": int((dfp["RECEIVE MATCH"] == "NO").sum()) if "RECEIVE MATCH" in dfp.columns else None,
                "PCS MATCH NO": int((dfp["PCS MATCH"] == "NO").sum()) if "PCS MATCH" in dfp.columns else None,
                "SKU MATCH NO": int((dfp["MATCH"] == "NO").sum()) if "MATCH" in dfp.columns else None,
            }
        )

        st.dataframe(dfp.head(100), use_container_width=True)

        st.download_button(
            "⬇️ Download SKU + Receive + PCS Match Excel",
            to_excel_bytes(dfp),
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
