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
    """
    Normalize IDs like:
      - 7052583281.0 -> 7052583281
      - '7052583281  -> 7052583281
      - 7.05258E+09  -> 7052583281
      - CHS2020      -> CHS2020
    """
    if pd.isna(x):
        return ""
    s = str(x).strip()
    if s.startswith("'"):
        s = s[1:].strip()
    s = s.replace(",", "")

    # float/scientific handling
    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except:
        pass

    return s


def norm_int_str(x) -> str:
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


def series_digits_only(s: pd.Series) -> bool:
    s = s.astype(str).str.strip().replace({"": pd.NA}).dropna()
    if len(s) == 0:
        return True
    return s.str.fullmatch(r"\d+").all()


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
    return "DOG EAR"  # default


def sku_is_valid(val) -> bool:
    if pd.isna(val):
        return False
    v = str(val).strip().upper()
    return v not in ("", "NAN", "NONE")


def load_sku_lookup(sku_file) -> pd.DataFrame:
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
# PDF Header extraction
# ==================================================
def extract_container_and_order(full_text: str, filename: str):
    # container like CAAU5869568 / FANU1196717
    container = ""
    m = re.search(r"\b([A-Z]{4}\d{7})\b", full_text)
    if m:
        container = m.group(1)
    else:
        m2 = re.search(r"\b([A-Z]{4}\d{7})\b", (filename or "").upper())
        if m2:
            container = m2.group(1)

    # order like 77660-2 / 81804-2 (from around "P.O.")
    order = ""
    lines = full_text.splitlines()
    po_idx = None
    for i, line in enumerate(lines):
        up = line.upper()
        if "P.O." in up or "PO #" in up or up.strip() == "P.O. #:":
            po_idx = i
            break

    if po_idx is not None:
        for j in range(po_idx, min(po_idx + 25, len(lines))):
            for tok in re.split(r"\s+", lines[j].replace(":", " ")):
                if re.fullmatch(r"\d{5,}(?:-\d+)?", tok.strip()):
                    order = tok.strip()
                    break
            if order:
                break

    return container, order


# ==================================================
# PDF Line Item Parser (handles multiple formats)
# ==================================================
def parse_pdfs_line_items(pdf_files, package_id_whitelist=None) -> pd.DataFrame:
    """
    Robust parser that works when:
      - LPN is numeric (8–12 digits) OR alphanumeric (CHS2020)
      - PIECES can appear before or after the LPN
      - Some PDFs list: GRADE DIM LPN PCS ...
      - Others list: GRADE DIM PCS LPN ...
    Uses DIM + PIECES as anchors, then selects the best identifier near PIECES.

    Returns columns:
      PACKAGEID, PCS, QTY, GRADE, THICKNESS, WIDTH, LENGTH, CONTAINER, ORDERNUMBER, PDF_FILE
    """
    dim_pat = re.compile(r"^\d+\s*[Xx]\s*\d+\s*[Xx]\s*\d+$")
    int_pat = re.compile(r"^\d+$")

    def is_pieces(tok: str) -> bool:
        return bool(int_pat.fullmatch(tok)) and (1 <= int(tok) <= 5000)

    def id_pattern_score(tok: str) -> int:
        # ignore decimals
        if "." in tok:
            return 0
        if dim_pat.match(tok):
            return 0
        if is_pieces(tok):
            return 0
        if re.fullmatch(r"\d{8,12}", tok):
            return 2
        if re.fullmatch(r"[A-Za-z]+[0-9]+[A-Za-z0-9]*", tok):
            return 1
        return 0

    rows = []

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        pages_text = [(p.extract_text() or "") for p in reader.pages]
        full_text = "\n".join(pages_text)

        container, order = extract_container_and_order(full_text, pdf.name)

        for text in pages_text:
            for line in text.splitlines():
                tokens = [t for t in line.strip().split() if t]
                if len(tokens) < 4:
                    continue

                # find DIM token
                dim_idx = None
                for i, tok in enumerate(tokens):
                    if dim_pat.match(tok):
                        dim_idx = i
                        break
                if dim_idx is None:
                    continue

                # parse dimensions
                dims = re.sub(r"\s+", "", tokens[dim_idx])
                parts = re.split(r"[Xx]", dims)
                if len(parts) != 3:
                    continue
                try:
                    thk = int(parts[0]); wid = int(parts[1]); leng = int(parts[2])
                except:
                    continue

                # grade can be multi-token (DOG EAR), so join everything before DIM
                grade = " ".join(tokens[:dim_idx]).strip()

                # find PIECES near/after DIM
                pieces_idx = None
                for j in range(dim_idx + 1, min(dim_idx + 9, len(tokens))):
                    if is_pieces(tokens[j]):
                        pieces_idx = j
                        break
                if pieces_idx is None:
                    continue

                pieces = int(tokens[pieces_idx])

                # find best PACKAGEID token near PIECES
                candidates = []
                for off in range(-6, 7):
                    if off == 0:
                        continue
                    k = pieces_idx + off
                    if 0 <= k < len(tokens):
                        tok = tokens[k]
                        score = id_pattern_score(tok)
                        if score <= 0:
                            continue

                        bonus = 0
                        if package_id_whitelist is not None and tok in package_id_whitelist:
                            bonus = 100  # force-match if it exists in container list

                        candidates.append((tok, bonus + score, abs(off)))

                if not candidates:
                    continue

                # best: highest score, then closest distance
                candidates.sort(key=lambda x: (-x[1], x[2]))
                package_id = candidates[0][0]

                # compute BF quantity
                qty = int(round(pieces * (thk * wid * leng) / 144.0))

                rows.append({
                    "PACKAGEID": package_id,
                    "PCS": pieces,
                    "QTY": qty,
                    "GRADE": grade,
                    "THICKNESS": str(thk),
                    "WIDTH": str(wid),
                    "LENGTH": str(leng),
                    "CONTAINER": container,
                    "ORDERNUMBER": order,
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

    # normalize container ids
    df["PACKAGEID"] = df["PACKAGEID"].apply(norm_id)
    df["PCS"] = df["PCS"].apply(norm_int_str)

    package_whitelist = set(df["PACKAGEID"].astype(str))

    # parse PDFs once into line-items (supports both formats)
    pdf_items = parse_pdfs_line_items(pdf_files, package_id_whitelist=package_whitelist)

    pdf_lpns = set(pdf_items["PACKAGEID"].astype(str)) if not pdf_items.empty else set()
    pcs_map = {}
    if not pdf_items.empty:
        pcs_map = (
            pdf_items.groupby("PACKAGEID")["PCS"]
            .first()
            .to_dict()
        )

    df["PDF LPN"] = df["PACKAGEID"].astype(str).apply(lambda x: x if x in pdf_lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].astype(str).apply(lambda x: "YES" if x in pdf_lpns else "NO")

    df["PCS CHECK"] = df["PACKAGEID"].astype(str).apply(lambda x: str(pcs_map.get(x, "")))

    def pcs_match(container_pcs, pdf_pcs):
        try:
            return "YES" if int(container_pcs) == int(pdf_pcs) else "NO"
        except:
            return "NO"

    df["PCS MATCH"] = df.apply(lambda r: pcs_match(r.get("PCS", ""), r.get("PCS CHECK", "")), axis=1)

    # SKU match
    sku_df = load_sku_lookup(sku_file)
    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)
    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|" +
        df["THICKNESS"].astype(str) + "|" +
        df["WIDTH"].astype(str) + "|" +
        df["LENGTH"].astype(str)
    )

    df = df.merge(sku_df[["SKU", "MATCH KEY"]], how="left", on="MATCH KEY")
    df["MATCH"] = df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")
    df = df.fillna("")

    # audit column ordering
    audit_cols = ["PDF LPN", "RECEIVE MATCH", "PCS CHECK", "PCS MATCH", "SKU", "MATCH"]
    existing_audit = [c for c in audit_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_audit]
    df = df[others + existing_audit]

    return df


def fix_pcs_mismatch_use_container_truth(df: pd.DataFrame):
    """
    Container list PCS is truth.
    Fix ONLY the audit columns:
      PCS CHECK <- PCS (where PCS CHECK exists and differs)
      PCS MATCH <- YES
    Do NOT modify PCS.
    """
    if df is None or df.empty:
        return df, 0

    needed = {"PCS", "PCS CHECK", "PCS MATCH"}
    if not needed.issubset(df.columns):
        return df, 0

    out = df.copy()

    pcs = out["PCS"].astype(str).str.strip()
    pcs_check = out["PCS CHECK"].astype(str).str.strip()
    pcs_match_col = out["PCS MATCH"].astype(str).str.upper()

    mask = (pcs != "") & (pcs_check != "") & (pcs != pcs_check) & (pcs_match_col == "NO")

    changed = int(mask.sum())
    if changed > 0:
        out.loc[mask, "PCS CHECK"] = out.loc[mask, "PCS"]
        out.loc[mask, "PCS MATCH"] = "YES"

    return out, changed


# ==================================================
# Sales Assist Generator (supports numeric OR alphanumeric identifiers)
# ==================================================
def generate_sales_assist(df: pd.DataFrame) -> pd.DataFrame:
    order_str = df.get("ORDERNUMBER", "").astype(str).str.split("-").str[0].str.strip()
    if series_digits_only(order_str):
        order_out = pd.to_numeric(order_str, errors="coerce").fillna(0).astype(int)
    else:
        order_out = order_str

    ident_str = df.get("PACKAGEID", "").astype(str).str.strip()
    if series_digits_only(ident_str):
        ident_out = pd.to_numeric(ident_str, errors="coerce").fillna(0).astype(int)
    else:
        ident_out = ident_str

    return pd.DataFrame({
        "SKU": df.get("SKU", ""),
        "Pieces": pd.to_numeric(df.get("PCS", 0), errors="coerce").fillna(0).astype(int),
        "Quantity": pd.to_numeric(df.get("QTY", 0), errors="coerce").fillna(0).astype(int),
        "QuantityUOM": "BF",
        "PriceUOM": "MBF",
        "PricePerUOM": 0,
        "OrderNumber": order_out,
        "ContainerNumber": df.get("CONTAINER", ""),
        "ReloadReference": "",
        "Identifier": ident_out,
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

tab1, tab2 = st.tabs(["Full Match + Audit (Container + PDFs)", "PDF + SKU Lookup → Sales Assist (no container)"])


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

            try:
                st.dataframe(dfp.style.apply(highlight_mismatches, axis=1), use_container_width=True)
            except Exception:
                st.dataframe(dfp, use_container_width=True)

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
            items_df = parse_pdfs_line_items(pdf_files, package_id_whitelist=None)

            if items_df.empty:
                st.error(
                    "No line-items were parsed from the PDFs.\n\n"
                    "If these are scanned images, OCR would be required.\n"
                    "If they are text-based, the line format may be totally different from the supported patterns."
                )
            else:
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

                st.session_state.pdf_items_df = items_df
                st.session_state.pdf_sa_df = generate_sales_assist(items_df)

                st.success(f"Built Sales Assist from PDFs ({len(items_df)} line-items).")

        if st.session_state.pdf_items_df is not None:
            st.write("Parsed PDF line-items preview:")
            st.dataframe(st.session_state.pdf_items_df.head(200), use_container_width=True)

        if st.session_state.pdf_sa_df is not None:
            sa_name_pdf = st.text_input("Sales Assist file name (no extension)", value="Sales_Assist_From_PDFs")
            st.download_button(
                "⬇️ Download Sales Assist Excel (PDF-only)",
                to_excel_bytes(st.session_state.pdf_sa_df),
                f"{sa_name_pdf}.xlsx"
            )
