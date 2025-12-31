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
if "pdf_items_df" not in st.session_state:
    st.session_state.pdf_items_df = None
if "pdf_sa_df" not in st.session_state:
    st.session_state.pdf_sa_df = None


# ==================================================
# Helpers
# ==================================================
def to_excel_bytes(df):
    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    return bio


def normalize_headers(raw):
    header_row = None
    for i in range(min(20, len(raw))):
        if "GRADE" in raw.iloc[i].astype(str).str.upper().values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not locate header row containing GRADE")

    raw.columns = raw.iloc[header_row].astype(str).str.strip().str.upper()
    return raw.iloc[header_row + 1:].reset_index(drop=True)


def norm_id(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().replace(",", "")
    m = re.search(r"[A-Z0-9]{4,12}", s)
    return m.group(0) if m else s


def norm_int(x):
    try:
        return int(float(str(x)))
    except:
        return ""


# ==================================================
# SKU Logic
# ==================================================
def map_description(grade):
    g = str(grade).upper()
    if "APG" in g:
        return "TAEDA PINE APG"
    if "DOG" in g:
        return "DOG EAR"
    if "III" in g or "3COM" in g:
        return "TAEDA PINE #3 COMMON"
    return "DOG EAR"


def sku_is_valid(v):
    if pd.isna(v):
        return False
    return str(v).strip() not in ("", "NAN", "NONE")


def load_sku_lookup(sku_file):
    REQUIRED = {
        "SKU": ["SKU"],
        "DESCRIPTION": ["DESCRIPTION", "DESC"],
        "THICKNESS": ["THICKNESS", "THK"],
        "WIDTH": ["WIDTH", "W"],
        "LENGTH": ["LENGTH", "LEN"],
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
# PDF PARSER (MATCHES YOUR PDFs EXACTLY)
# ==================================================
def parse_pdfs_line_items(pdf_files):
    dim_pat = re.compile(r"^\d+[xX]\d+[xX]\d+$")
    rows = []

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        text = "\n".join(p.extract_text() or "" for p in reader.pages)

        # header info
        container = re.search(r"\b[A-Z]{4}\d{7}\b", text)
        container = container.group(0) if container else ""

        order = re.search(r"\b\d{5,}-\d+\b", text)
        order = order.group(0) if order else ""

        for line in text.splitlines():
            tokens = line.split()
            if len(tokens) < 5:
                continue

            for i, tok in enumerate(tokens):
                if not dim_pat.match(tok):
                    continue

                grade = " ".join(tokens[:i])
                dim = tok.lower().replace("x", "X")
                lpn = tokens[i + 1]
                pcs = tokens[i + 2]

                if not pcs.isdigit():
                    continue

                t, w, l = [int(x) for x in dim.split("X")]
                bf = round(int(pcs) * (t * w * l) / 144)

                rows.append({
                    "PACKAGEID": lpn,
                    "PCS": int(pcs),
                    "QTY": bf,
                    "GRADE": grade,
                    "THICKNESS": str(t),
                    "WIDTH": str(w),
                    "LENGTH": str(l),
                    "CONTAINER": container,
                    "ORDERNUMBER": order,
                    "PDF_FILE": pdf.name
                })

    return pd.DataFrame(rows)


# ==================================================
# FIX PCS MISMATCH (container list is truth)
# ==================================================
def fix_pcs_mismatch_use_container_truth(df):
    mask = (
        (df["PCS"].astype(str) != df["PCS CHECK"].astype(str)) &
        (df["PCS"].astype(str) != "") &
        (df["PCS CHECK"].astype(str) != "")
    )
    df.loc[mask, "PCS CHECK"] = df.loc[mask, "PCS"]
    df.loc[mask, "PCS MATCH"] = "YES"
    return df, int(mask.sum())


# ==================================================
# SALES ASSIST
# ==================================================
def generate_sales_assist(df):
    return pd.DataFrame({
        "SKU": df["SKU"],
        "Pieces": df["PCS"],
        "Quantity": df["QTY"],
        "QuantityUOM": "BF",
        "PriceUOM": "MBF",
        "PricePerUOM": 0,
        "OrderNumber": df["ORDERNUMBER"].astype(str).str.split("-").str[0],
        "ContainerNumber": df["CONTAINER"],
        "ReloadReference": "",
        "Identifier": df["PACKAGEID"],
        "ProFormaPrice": 0
    })


# ==================================================
# STREAMLIT UI
# ==================================================
st.set_page_config(page_title="BIFP Import Checker", layout="wide")
st.title("📦 BIFP PDF → SKU → Sales Assist")

sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

if sku_file and pdf_files:
    if st.button("Parse PDFs + Build Sales Assist"):
        items = parse_pdfs_line_items(pdf_files)

        if items.empty:
            st.error("No line-items parsed from PDFs.")
        else:
            sku_df = load_sku_lookup(sku_file)
            items["MAPPED DESCRIPTION"] = items["GRADE"].apply(map_description)
            items["MATCH KEY"] = (
                items["MAPPED DESCRIPTION"] + "|" +
                items["THICKNESS"] + "|" +
                items["WIDTH"] + "|" +
                items["LENGTH"]
            )

            items = items.merge(
                sku_df[["SKU", "MATCH KEY"]],
                how="left",
                on="MATCH KEY"
            )
            items["MATCH"] = items["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")

            st.session_state.pdf_items_df = items
            st.session_state.pdf_sa_df = generate_sales_assist(items)

            st.success(f"Parsed {len(items)} line-items from PDFs.")
            st.dataframe(items, use_container_width=True)

            st.download_button(
                "⬇️ Download Sales Assist Excel",
                to_excel_bytes(st.session_state.pdf_sa_df),
                "Sales_Assist_From_PDFs.xlsx"
            )
