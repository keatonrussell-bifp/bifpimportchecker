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
def to_excel_bytes(df: pd.DataFrame) -> BytesIO:
    bio = BytesIO()
    df.to_excel(bio, index=False)
    bio.seek(0)
    return bio


def normalize_headers(raw: pd.DataFrame) -> pd.DataFrame:
    header_row = None
    for i in range(min(20, len(raw))):
        if "GRADE" in raw.iloc[i].astype(str).str.upper().values:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not locate header row containing GRADE")

    raw.columns = raw.iloc[header_row].astype(str).str.strip().str.upper()
    return raw.iloc[header_row + 1:].reset_index(drop=True)


def norm_id(x) -> str:
    """Normalize IDs like 7052583281.0, '7052583281, scientific notation, etc."""
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.startswith("'"):
        s = s[1:].strip()
    s = s.replace(",", "")

    try:
        f = float(s)
        if f.is_integer():
            s = str(int(f))
    except:
        pass

    m = re.search(r"\d{8,12}", s)
    return m.group(0) if m else s


def norm_int_str(x) -> str:
    """Normalize integers; blank if not parseable."""
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


# ==================================================
# SKU Logic
# ==================================================
def map_description(grade) -> str:
    grade = str(grade).upper()
    if "APG" in grade:
        return "TAEDA PINE APG"
    if "DOG" in grade:
        return "DOG EAR"
    if re.search(r"\bIII/V\b|\bIII\b|\b3COM\b", grade):
        return "TAEDA PINE #3 COMMON"
    # per your request: blank/unknown defaults to DOG EAR
    return "DOG EAR"


def sku_is_valid(val) -> bool:
    if pd.isna(val):
        return False
    v = str(val).strip().upper()
    return v not in ("", "NAN", "NONE")


def load_sku_lookup(sku_file) -> pd.DataFrame:
    """
    Load SKU lookup from any sheet that contains required columns (or common aliases).
    Normalizes to: SKU, DESCRIPTION, THICKNESS, WIDTH, LENGTH
    """
    REQUIRED = {
        "SKU": ["SKU"],
        "DESCRIPTION": ["DESCRIPTION", "DESC", "PRODUCT DESCRIPTION", "GRADE DESC"],
        "THICKNESS": ["THICKNESS", "THK", "THICK"],
        "WIDTH": ["WIDTH", "W"],
        "LENGTH": ["LENGTH", "LEN", "L"],
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


# ==================================================
# PDF Parsing (LPN + PIECES adjacency)
# ==================================================
def extract_lpns_and_pieces_from_pdfs(pdf_files):
    """
    Returns:
      lpns_set: set of LPNs (8-12 digits)
      pieces_map: {LPN: PIECES}

    Works for extracted lines like:
      3COM 1X4X96 240 7052583281 1120.0000
      3COM 1X4X96 7052583281 240 1120.0000
    """
    lpns_set = set()
    pieces_map = {}

    lpn_findall = re.compile(r"\b\d{8,12}\b")
    lpn_token = re.compile(r"^\d{8,12}$")
    int_token = re.compile(r"^\d+$")

    def is_reasonable_pieces(tok: str) -> bool:
        if not int_token.match(tok):
            return False
        v = int(tok)
        return 1 <= v <= 5000

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        for page in reader.pages:
            text = page.extract_text()
            if not text:
                continue

            lpns_set.update(lpn_findall.findall(text))

            for line in text.splitlines():
                tokens = line.strip().split()
                if not tokens:
                    continue

                for idx, tok in enumerate(tokens):
                    if not lpn_token.match(tok):
                        continue

                    lpn = tok
                    pieces_val = None

                    for off in (-1, 1, -2, 2, -3, 3):
                        j = idx + off
                        if 0 <= j < len(tokens) and is_reasonable_pieces(tokens[j]):
                            pieces_val = tokens[j]
                            break

                    if pieces_val and lpn not in pieces_map:
                        pieces_map[lpn] = pieces_val

    return lpns_set, pieces_map


def extract_pdf_header_info(pages_text, pdf_filename: str):
    """
    Extract container + PO/order number from header-ish text.
    Avoids false matches against LPNs.
    """
    # container id like CAAU5869568
    container_pat = re.compile(r"\b([A-Z]{4}\d{7})\b")

    # order number like 77660-2 or 77660
    order_pat = re.compile(r"^\d{5,}(?:-\d+)?$")

    full_text = "\n".join(pages_text)
    container = ""
    m = container_pat.search(full_text)
    if m:
        container = m.group(1)
    else:
        m2 = container_pat.search(pdf_filename.upper())
        if m2:
            container = m2.group(1)

    # find P.O. # value by scanning lines after a line containing "P.O."
    ordernum = ""
    lines = full_text.splitlines()
    po_idx = None
    for i, line in enumerate(lines):
        if "P.O." in line.upper():
            po_idx = i
            break

    if po_idx is not None:
        for j in range(po_idx, min(po_idx + 15, len(lines))):
            for tok in lines[j].replace(":", " ").split():
                tok = tok.strip()
                if order_pat.match(tok):
                    ordernum = tok
                    break
            if ordernum:
                break

    return container, ordernum


def parse_pdfs_to_items_df(pdf_files):
    """
    Build a row-per-LPN dataframe from PDFs:
      PACKAGEID (LPN), PCS (PIECES), QTY (BF), GRADE, THICKNESS, WIDTH, LENGTH, CONTAINER, ORDERNUMBER
    """
    dim_pat = re.compile(r"^\d+\s*[Xx]\s*\d+\s*[Xx]\s*\d+$")
    lpn_token = re.compile(r"^\d{8,12}$")
    int_token = re.compile(r"^\d+$")

    def is_reasonable_pieces(tok: str) -> bool:
        if not int_token.match(tok):
            return False
        v = int(tok)
        return 1 <= v <= 5000

    rows = []

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        pages_text = [(p.extract_text() or "") for p in reader.pages]

        container, ordernum = extract_pdf_header_info(pages_text, pdf.name)

        for text in pages_text:
            for line in text.splitlines():
                tokens = [t.strip() for t in line.split() if t.strip()]
                if len(tokens) < 4:
                    continue

                # find dimension token
                dim_idx = None
                for i, t in enumerate(tokens):
                    if dim_pat.match(t):
                        dim_idx = i
                        break
                if dim_idx is None:
                    continue

                grade_tok = tokens[dim_idx - 1] if dim_idx > 0 else ""
                dims = re.sub(r"\s+", "", tokens[dim_idx]).upper()
                parts = re.split(r"[Xx]", dims)
                if len(parts) != 3:
                    continue

                try:
                    thk = int(parts[0])
                    wid = int(parts[1])
                    leng = int(parts[2])
                except:
                    continue

                # find LPN and adjacent pieces
                lpn_idx = None
                for i, t in enumerate(tokens):
                    if lpn_token.match(t):
                        lpn_idx = i
                        break
                if lpn_idx is None:
                    continue

                lpn = tokens[lpn_idx]
                pieces = None
                for off in (-1, 1, -2, 2, -3, 3):
                    j = lpn_idx + off
                    if 0 <= j < len(tokens) and is_reasonable_pieces(tokens[j]):
                        pieces = int(tokens[j])
                        break
                if pieces is None:
                    continue

                # compute BF quantity
                bf = pieces * (thk * wid * leng) / 144.0
                qty = int(round(bf))

                rows.append({
                    "PACKAGEID": lpn,
                    "PCS": pieces,
                    "QTY": qty,
                    "GRADE": grade_tok,
                    "THICKNESS": str(thk),
                    "WIDTH": str(wid),
                    "LENGTH": str(leng),
                    "CONTAINER": container,
                    "ORDERNUMBER": ordernum,
                    "PDF_FILE": pdf.name
                })

    return pd.DataFrame(rows)


# ==================================================
# Full Process (Container + PDFs + SKU)
# ==================================================
def process_all(container_file, sku_file, pdf_files):
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    required_cols = {"PACKAGEID", "PCS", "GRADE", "THICKNESS", "WIDTH", "LENGTH"}
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Container list missing required columns: {missing}")

    df["PACKAGEID"] = df["PACKAGEID"].apply(norm_id)
    df["PCS"] = df["PCS"].apply(norm_int_str)

    lpns, pieces_map = extract_lpns_and_pieces_from_pdfs(pdf_files)

    df["PDF LPN"] = df["PACKAGEID"].apply(lambda x: x if x in lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].apply(lambda x: "YES" if x in lpns else "NO")

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
    df = df.fillna("")

    # audit-friendly ordering
    audit_cols = ["PDF LPN", "RECEIVE MATCH", "PCS CHECK", "PCS MATCH", "SKU", "MATCH"]
    existing_audit = [c for c in audit_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_audit]
    return df[others + existing_audit]


def fix_pcs_mismatch_use_container_truth(df: pd.DataFrame):
    """
    Your clarified behavior:
      - Container list PCS is truth
      - Only update the mismatch fields in the match excel:
          PCS CHECK <- PCS   (where PCS CHECK exists and differs)
          PCS MATCH <- YES
      - Do NOT modify PCS
    """
    if df is None or df.empty:
        return df, 0

    needed = {"PCS", "PCS CHECK", "PCS MATCH"}
    if not needed.issubset(set(df.columns)):
        return df, 0

    out = df.copy()

    pcs = out["PCS"].astype(str).str.strip()
    pcs_check = out["PCS CHECK"].astype(str).str.strip()

    mask = (
        (pcs != "") &
        (pcs_check != "") &
        (out["PCS MATCH"].astype(str).str.upper() == "NO") &
        (pcs != pcs_check)
    )

    changed = int(mask.sum())
    if changed > 0:
        out.loc[mask, "PCS CHECK"] = out.loc[mask, "PCS"]
        out.loc[mask, "PCS MATCH"] = "YES"

    return out, changed


# ==================================================
# Sales Assist Generator (reused)
# ==================================================
def generate_sales_assist(df: pd.DataFrame) -> pd.DataFrame:
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


# ==================================================
# UI Styling
# ==================================================
def highlight_mismatches(row):
    if (
        row.get("RECEIVE MATCH") != "YES"
        or row.get("PCS MATCH") != "YES"
        or row.get("MATCH") != "YES"
    ):
        return ["background-color: #ffcccc"] * len(row)
    return [""] * len(row)


# ==================================================
# Streamlit UI
# ==================================================
st.set_page_config(page_title="BIFP Import Checker", layout="wide")
st.title("📦 BIFP SKU + Receive + PCS Match + Sales Assist")

container_file = st.file_uploader("Upload Container List Excel (optional)", type="xlsx")
sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

tab1, tab2 = st.tabs(["Full Match + Audit (Container + PDFs)", "PDF + SKU Lookup → Sales Assist"])


# --------------------------------------------------
# TAB 1: Full process
# --------------------------------------------------
with tab1:
    st.subheader("Full Match + Audit")

    if not (container_file and sku_file and pdf_files):
        st.info("Upload **Container List + SKU Lookup + PDFs** to run the full match.")
    else:
        if st.button("Run Full Process"):
            st.session_state.processed_df = process_all(container_file, sku_file, pdf_files)
            st.success("Full process completed")

        if st.session_state.processed_df is not None:
            dfp = st.session_state.processed_df

            # Preview with red mismatches
            try:
                st.dataframe(dfp.style.apply(highlight_mismatches, axis=1), use_container_width=True)
            except Exception:
                st.dataframe(dfp, use_container_width=True)

            # Download + Fix buttons side-by-side
            c1, c2 = st.columns([1, 1])

            with c1:
                st.download_button(
                    "⬇️ Download Match Excel",
                    to_excel_bytes(dfp),
                    container_file.name.replace(".xlsx", "_SKU_RECEIVE_PCS_MATCH.xlsx"),
                )

            with c2:
                if st.button("FIX PCS Mismatch"):
                    fixed, n = fix_pcs_mismatch_use_container_truth(dfp)
                    st.session_state.processed_df = fixed
                    if n == 0:
                        st.info("No PCS mismatches found to fix (or PCS CHECK was blank).")
                    else:
                        st.success(f"Fixed {n} PCS mismatches using **container PCS as truth**.")

            st.divider()

            # Sales Assist from FULL process
            st.subheader("Sales Assist Export (from Full Match)")

            sa_name = st.text_input("Sales Assist file name (no extension)", value="Sales_Assist_Full")

            if st.button("Generate Sales Assist Excel (Full Match)"):
                sa_df = generate_sales_assist(st.session_state.processed_df)
                st.download_button(
                    "⬇️ Download Sales Assist Excel",
                    to_excel_bytes(sa_df),
                    f"{sa_name}.xlsx"
                )


# --------------------------------------------------
# TAB 2: PDF-only Sales Assist
# --------------------------------------------------
with tab2:
    st.subheader("PDF + SKU Lookup → Sales Assist (no container list)")

    if not (sku_file and pdf_files):
        st.info("Upload **SKU Lookup + PDFs** to generate Sales Assist directly from PDFs.")
    else:
        if st.button("Parse PDFs + Match SKU + Build Sales Assist"):
            # Parse PDFs into item rows
            items_df = parse_pdfs_to_items_df(pdf_files)

            if items_df.empty:
                st.error("No line-items were parsed from the PDFs. (If these are scanned images, OCR would be required.)")
            else:
                # SKU match
                sku_df = load_sku_lookup(sku_file)
                items_df["MAPPED DESCRIPTION"] = items_df["GRADE"].apply(map_description)
                items_df["MATCH KEY"] = (
                    items_df["MAPPED DESCRIPTION"] + "|" +
                    items_df["THICKNESS"].astype(str) + "|" +
                    items_df["WIDTH"].astype(str) + "|" +
                    items_df["LENGTH"].astype(str)
                )

                items_df = items_df.merge(
                    sku_df[["SKU", "MATCH KEY"]],
                    how="left",
                    on="MATCH KEY"
                )
                items_df["MATCH"] = items_df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")

                # Save in state
                st.session_state.pdf_items_df = items_df
                st.session_state.pdf_sa_df = generate_sales_assist(items_df)

                st.success(f"Built Sales Assist from PDFs ({len(items_df)} LPN rows).")

        if st.session_state.pdf_items_df is not None:
            st.write("Parsed PDF rows preview:")
            st.dataframe(st.session_state.pdf_items_df.head(100), use_container_width=True)

        if st.session_state.pdf_sa_df is not None:
            sa_name_pdf = st.text_input("Sales Assist file name (no extension)", value="Sales_Assist_From_PDFs")
            st.download_button(
                "⬇️ Download Sales Assist Excel (PDF-only)",
                to_excel_bytes(st.session_state.pdf_sa_df),
                f"{sa_name_pdf}.xlsx"
            )
