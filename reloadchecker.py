import streamlit as st
import pandas as pd
import re
from PyPDF2 import PdfReader
from io import BytesIO
from collections import defaultdict, deque


# ==================================================
# Session State
# ==================================================
if "processed_df" not in st.session_state:
    st.session_state.processed_df = None

if "pdf_items_df" not in st.session_state:
    st.session_state.pdf_items_df = None

if "pdf_sa_df" not in st.session_state:
    st.session_state.pdf_sa_df = None

if "dabg_df" not in st.session_state:
    st.session_state.dabg_df = None

if "dabg_pdf_df" not in st.session_state:
    st.session_state.dabg_pdf_df = None

if "dabg_sa_df" not in st.session_state:
    st.session_state.dabg_sa_df = None


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
        vals = raw.iloc[i].astype(str).str.upper().str.strip().values
        if "GRADE" in vals:
            header_row = i
            break

    if header_row is None:
        raise ValueError("Could not locate header row containing GRADE")

    raw.columns = raw.iloc[header_row].astype(str).str.strip().str.upper()
    return raw.iloc[header_row + 1:].reset_index(drop=True)


def norm_id(x) -> str:
    if pd.isna(x):
        return ""

    s = str(x).strip()

    if s.startswith("'"):
        s = s[1:].strip()

    s = s.replace(",", "")

    try:
        f = float(s)
        if f.is_integer():
            return str(int(f))
    except Exception:
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
    except Exception:
        return ""


def series_digits_only(s: pd.Series) -> bool:
    s = s.astype(str).str.strip().replace({"": pd.NA}).dropna()

    if len(s) == 0:
        return True

    return s.str.fullmatch(r"\d+").all()


def col_or_default(df: pd.DataFrame, col: str, default):
    if df is None:
        return pd.Series([], dtype="object")

    if col in df.columns:
        return df[col]

    return pd.Series([default] * len(df), index=df.index)


# ==================================================
# SKU Logic
# ==================================================
def map_description(grade) -> str:
    g = "" if pd.isna(grade) else str(grade).strip().upper()

    if g == "":
        return "DOG EAR"

    if g == "FB":
        return "DOG EAR"

    if re.search(r"\bD[-\s]?GRADE\b|\bGS[-\s]?D[-\s]?GRADE\b|\bGRADE[-\s]?D\b", g):
        return "TAEDA PINE D"

    if "APG" in g:
        return "TAEDA PINE APG"

    if "DOG" in g:
        return "DOG EAR"

    if "FENCE" in g:
        return "DOG EAR"

    if re.search(r"\bIII/V\b|\bIII\b|\b3COM\b", g):
        return "TAEDA PINE #3 COMMON"

    return ""


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

            df["DESCRIPTION"] = df["DESCRIPTION"].astype(str).str.upper().str.strip()
            df["THICKNESS"] = df["THICKNESS"].apply(norm_int_str)
            df["WIDTH"] = df["WIDTH"].apply(norm_int_str)
            df["LENGTH"] = df["LENGTH"].apply(norm_int_str)

            df["MATCH KEY"] = (
                df["DESCRIPTION"] + "|"
                + df["THICKNESS"] + "|"
                + df["WIDTH"] + "|"
                + df["LENGTH"]
            )

            return df

    raise ValueError("SKU lookup missing required columns: SKU, DESCRIPTION, THICKNESS, WIDTH, LENGTH.")


# ==================================================
# PDF Header Extraction
# ==================================================
def extract_container_and_order(full_text: str, filename: str):
    container = ""

    m = re.search(r"\b([A-Z]{4}\d{7})\b", full_text)
    if m:
        container = m.group(1)
    else:
        m2 = re.search(r"\b([A-Z]{4}\d{7})\b", (filename or "").upper())
        if m2:
            container = m2.group(1)

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
# Existing PDF Line Item Parser
# ==================================================
def parse_pdfs_line_items(pdf_files, package_id_whitelist=None) -> pd.DataFrame:
    dim_pat = re.compile(r"^\d+\s*[Xx]\s*\d+\s*[Xx]\s*\d+$")
    int_pat = re.compile(r"^\d+$")

    def is_pieces(tok: str) -> bool:
        return bool(int_pat.fullmatch(tok)) and (1 <= int(tok) <= 5000)

    def id_pattern_score(tok: str) -> int:
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

                dim_idx = None

                for i, tok in enumerate(tokens):
                    if dim_pat.match(tok):
                        dim_idx = i
                        break

                if dim_idx is None:
                    continue

                dims = re.sub(r"\s+", "", tokens[dim_idx])
                parts = re.split(r"[Xx]", dims)

                if len(parts) != 3:
                    continue

                try:
                    thk = int(parts[0])
                    wid = int(parts[1])
                    leng = int(parts[2])
                except Exception:
                    continue

                grade = " ".join(tokens[:dim_idx]).strip()

                pieces_idx = None

                for j in range(dim_idx + 1, min(dim_idx + 9, len(tokens))):
                    if is_pieces(tokens[j]):
                        pieces_idx = j
                        break

                if pieces_idx is None:
                    continue

                pieces = int(tokens[pieces_idx])

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
                            bonus = 100

                        candidates.append((tok, bonus + score, abs(off)))

                if not candidates:
                    continue

                candidates.sort(key=lambda x: (-x[1], x[2]))
                package_id = candidates[0][0]

                qty = int(round(pieces * (thk * wid * leng) / 144.0))

                rows.append(
                    {
                        "PACKAGEID": package_id,
                        "PCS": pieces,
                        "QTY": qty,
                        "GRADE": grade,
                        "THICKNESS": str(thk),
                        "WIDTH": str(wid),
                        "LENGTH": str(leng),
                        "CONTAINER": container,
                        "ORDERNUMBER": order,
                        "PDF_FILE": pdf.name,
                    }
                )

    return pd.DataFrame(rows)


# ==================================================
# DABG Helpers
# ==================================================
def make_dabg_dim_pcs_key(thickness, width, length, pcs) -> str:
    """
    DABG matching ignores grade.

    Example:
      1, 4, 144, 240 -> 1|4|144|240
    """
    return (
        norm_int_str(thickness) + "|"
        + norm_int_str(width) + "|"
        + norm_int_str(length) + "|"
        + norm_int_str(pcs)
    )


def parse_dabg_pdfs_lpn_rows(pdf_files) -> pd.DataFrame:
    """
    Parses warehouse receipt PDFs into LPN rows for DABG.

    New DABG key ignores grade:
      DABG_MATCH_KEY = THICKNESS|WIDTH|LENGTH|PCS

    Example:
      1|4|144|240

    Full consumed key:
      DABG_MATCH_KEY_LPN = THICKNESS|WIDTH|LENGTH|PCS|LPN

    Example:
      1|4|144|240|208
    """
    rows = []

    receipt_line_pat = re.compile(
        r"""
        ^\s*
        (?P<grade>[A-Za-z0-9#/\-\s]+?)
        \s+
        (?P<thickness>\d+)
        \s*[xX]\s*
        (?P<width>\d+)
        \s*[xX]\s*
        (?P<length>\d+)
        .*?
        \s+
        (?P<lpn>[A-Za-z0-9\-]+)
        \s+
        (?P<pieces>\d+)
        \s+
        (?P<total_lbs>\d+(?:\.\d+)?)
        \s*$
        """,
        re.VERBOSE,
    )

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        pages_text = [(p.extract_text() or "") for p in reader.pages]
        full_text = "\n".join(pages_text)

        container, order = extract_container_and_order(full_text, pdf.name)

        for page_num, text in enumerate(pages_text, start=1):
            for line_num, line in enumerate(text.splitlines(), start=1):
                line_clean = " ".join(line.strip().split())

                if not line_clean:
                    continue

                m = receipt_line_pat.match(line_clean)

                if not m:
                    continue

                grade_raw = m.group("grade").strip()
                thk = norm_int_str(m.group("thickness"))
                wid = norm_int_str(m.group("width"))
                leng = norm_int_str(m.group("length"))
                pcs = norm_int_str(m.group("pieces"))
                lpn = norm_id(m.group("lpn"))

                if not lpn or not pcs:
                    continue

                match_key = make_dabg_dim_pcs_key(thk, wid, leng, pcs)

                rows.append(
                    {
                        "PDF_FILE": pdf.name,
                        "PDF_PAGE": page_num,
                        "PDF_LINE": line_num,
                        "PDF_GRADE_RAW": grade_raw,
                        "THICKNESS": thk,
                        "WIDTH": wid,
                        "LENGTH": leng,
                        "PCS": pcs,
                        "LPN": lpn,
                        "DABG_MATCH_KEY": match_key,
                        "DABG_MATCH_KEY_LPN": f"{match_key}|{lpn}",
                        "CONTAINER": container,
                        "ORDERNUMBER": order,
                    }
                )

    return pd.DataFrame(rows)


# ==================================================
# Full Process: Existing Container + PDFs + SKU
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

    package_whitelist = set(df["PACKAGEID"].astype(str))

    pdf_items = parse_pdfs_line_items(pdf_files, package_id_whitelist=package_whitelist)

    pdf_lpns = set(pdf_items["PACKAGEID"].astype(str)) if not pdf_items.empty else set()

    pcs_map = {}

    if not pdf_items.empty:
        pcs_map = pdf_items.groupby("PACKAGEID")["PCS"].first().to_dict()

    df["PDF LPN"] = df["PACKAGEID"].astype(str).apply(lambda x: x if x in pdf_lpns else "")
    df["RECEIVE MATCH"] = df["PACKAGEID"].astype(str).apply(lambda x: "YES" if x in pdf_lpns else "NO")

    df["PCS CHECK"] = df["PACKAGEID"].astype(str).apply(lambda x: str(pcs_map.get(x, "")))

    def pcs_match(container_pcs, pdf_pcs):
        try:
            return "YES" if int(container_pcs) == int(pdf_pcs) else "NO"
        except Exception:
            return "NO"

    df["PCS MATCH"] = df.apply(lambda r: pcs_match(r.get("PCS", ""), r.get("PCS CHECK", "")), axis=1)

    sku_df = load_sku_lookup(sku_file)

    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)

    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|"
        + df["THICKNESS"].astype(str).apply(norm_int_str) + "|"
        + df["WIDTH"].astype(str).apply(norm_int_str) + "|"
        + df["LENGTH"].astype(str).apply(norm_int_str)
    )

    df = df.merge(sku_df[["SKU", "MATCH KEY"]], how="left", on="MATCH KEY")

    df["MATCH"] = df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")
    df = df.fillna("")

    audit_cols = ["PDF LPN", "RECEIVE MATCH", "PCS CHECK", "PCS MATCH", "SKU", "MATCH"]
    existing_audit = [c for c in audit_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_audit]

    df = df[others + existing_audit]

    return df


def process_dabg(container_file, sku_file, pdf_files):
    """
    DABG process:
      1. Read container list.
      2. Build container key: THICKNESS|WIDTH|LENGTH|PCS.
      3. Parse warehouse receipt PDF LPN rows.
      4. Build PDF key: THICKNESS|WIDTH|LENGTH|PCS.
      5. Consume one LPN per matching container row.
      6. Never reuse LPNs.
      7. Never make up missing LPNs.
      8. Still perform SKU match using the existing grade-based SKU logic.
    """
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    required_cols = {"GRADE", "THICKNESS", "WIDTH", "LENGTH", "PCS"}
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise ValueError(f"Container list missing required DABG columns: {missing}")

    df["PCS"] = df["PCS"].apply(norm_int_str)
    df["THICKNESS"] = df["THICKNESS"].apply(norm_int_str)
    df["WIDTH"] = df["WIDTH"].apply(norm_int_str)
    df["LENGTH"] = df["LENGTH"].apply(norm_int_str)

    # New container-side DABG key.
    # No grade. Only dimensions + pieces.
    df["DABG_CONTAINER_MATCH_KEY"] = df.apply(
        lambda r: make_dabg_dim_pcs_key(
            r.get("THICKNESS", ""),
            r.get("WIDTH", ""),
            r.get("LENGTH", ""),
            r.get("PCS", ""),
        ),
        axis=1,
    )

    pdf_df = parse_dabg_pdfs_lpn_rows(pdf_files)

    lpn_pool = defaultdict(deque)

    if not pdf_df.empty:
        pdf_df = pdf_df.reset_index(drop=True)

        for _, r in pdf_df.iterrows():
            key = r["DABG_MATCH_KEY"]
            lpn = r["LPN"]
            lpn_pool[key].append(lpn)

    assigned_lpns = []
    source_keys = []
    pdf_match_keys = []

    for _, r in df.iterrows():
        key = r["DABG_CONTAINER_MATCH_KEY"]

        if lpn_pool[key]:
            lpn = lpn_pool[key].popleft()
            assigned_lpns.append(lpn)
            pdf_match_keys.append(key)
            source_keys.append(f"{key}|{lpn}")
        else:
            assigned_lpns.append("")
            pdf_match_keys.append("")
            source_keys.append("")

    df["DABG_MATCH_KEY"] = df["DABG_CONTAINER_MATCH_KEY"]
    df["DABG_PACKAGEID"] = assigned_lpns
    df["DABG_MATCH_KEY_LPN"] = source_keys
    df["DABG LPN MATCH"] = df["DABG_PACKAGEID"].apply(lambda x: "YES" if str(x).strip() else "NO")

    # SKU match stays the same. This still uses grade/description.
    sku_df = load_sku_lookup(sku_file)

    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)

    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|"
        + df["THICKNESS"].astype(str).apply(norm_int_str) + "|"
        + df["WIDTH"].astype(str).apply(norm_int_str) + "|"
        + df["LENGTH"].astype(str).apply(norm_int_str)
    )

    df = df.merge(
        sku_df[["SKU", "MATCH KEY"]],
        how="left",
        on="MATCH KEY",
    )

    df["MATCH"] = df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")
    df = df.fillna("")

    priority_cols = [
        "DABG_CONTAINER_MATCH_KEY",
        "DABG_MATCH_KEY",
        "DABG_PACKAGEID",
        "DABG_MATCH_KEY_LPN",
        "DABG LPN MATCH",
        "SKU",
        "MATCH",
    ]

    existing_priority = [c for c in priority_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_priority]

    df = df[others + existing_priority]

    return df, pdf_df


def fix_pcs_mismatch_use_container_truth(df: pd.DataFrame):
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
# Sales Assist Generator
# ==================================================
def generate_sales_assist(df: pd.DataFrame, identifier_col: str = "PACKAGEID") -> pd.DataFrame:
    order_raw = col_or_default(df, "ORDERNUMBER", "").astype(str).str.split("-").str[0].str.strip()

    if series_digits_only(order_raw):
        order_out = pd.to_numeric(order_raw, errors="coerce").fillna(0).astype(int)
    else:
        order_out = order_raw

    ident_raw = col_or_default(df, identifier_col, "").astype(str).str.strip()

    if series_digits_only(ident_raw):
        ident_out = pd.to_numeric(ident_raw, errors="coerce").fillna(0).astype(int)
    else:
        ident_out = ident_raw

    pcs = pd.to_numeric(col_or_default(df, "PCS", 0), errors="coerce").fillna(0).astype(int)

    if "QTY" in df.columns:
        qty = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)
    else:
        thk = pd.to_numeric(col_or_default(df, "THICKNESS", 0), errors="coerce").fillna(0)
        wid = pd.to_numeric(col_or_default(df, "WIDTH", 0), errors="coerce").fillna(0)
        leng = pd.to_numeric(col_or_default(df, "LENGTH", 0), errors="coerce").fillna(0)

        qty = (pcs * (thk * wid * leng) / 144.0).round().fillna(0).astype(int)

    return pd.DataFrame(
        {
            "SKU": col_or_default(df, "SKU", ""),
            "Pieces": pcs,
            "Quantity": qty,
            "QuantityUOM": "BF",
            "PriceUOM": "MBF",
            "PricePerUOM": 0,
            "OrderNumber": order_out,
            "ContainerNumber": col_or_default(df, "CONTAINER", ""),
            "ReloadReference": "",
            "Identifier": ident_out,
            "ProFormaPrice": 0,
        }
    )


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


def highlight_dabg_mismatches(row):
    if (
        row.get("DABG LPN MATCH") != "YES"
        or row.get("MATCH") != "YES"
    ):
        return ["background-color: #ffcccc"] * len(row)

    return [""] * len(row)


# ==================================================
# Streamlit UI
# ==================================================
st.set_page_config(page_title="BIFP Import Checker", layout="wide")
st.title("📦 BIFP SKU + Receive + PCS Match + Sales Assist")

container_file = st.file_uploader("Upload Container List Excel optional", type="xlsx")
sku_file = st.file_uploader("Upload SKU Lookup Excel", type="xlsx")
pdf_files = st.file_uploader("Upload PDF Files", type="pdf", accept_multiple_files=True)

tab1, tab2, tab3 = st.tabs(
    [
        "Full Match + Audit Container + PDFs",
        "PDF + SKU Lookup → Sales Assist no container",
        "DABG",
    ]
)


# --------------------------------------------------
# TAB 1: Full process
# --------------------------------------------------
with tab1:
    st.subheader("Full Match + Audit")

    if not (container_file and sku_file and pdf_files):
        st.info("Upload Container List + SKU Lookup + PDFs to run the full match.")
    else:
        if st.button("Run Full Process"):
            st.session_state.processed_df = process_all(container_file, sku_file, pdf_files)
            st.success("Full process completed.")

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
                        st.info("No PCS mismatches found to fix, or PCS CHECK was blank.")
                    else:
                        st.success(f"Fixed {n} PCS mismatches using container PCS as truth.")

            st.divider()

            st.subheader("Sales Assist Export from Full Match")

            sa_name = st.text_input(
                "Sales Assist file name no extension",
                value="Sales_Assist_Full",
                key="sa_name_full",
            )

            if st.button("Generate Sales Assist Excel Full Match"):
                sa_df = generate_sales_assist(st.session_state.processed_df)

                st.download_button(
                    "⬇️ Download Sales Assist Excel",
                    to_excel_bytes(sa_df),
                    f"{sa_name}.xlsx",
                )


# --------------------------------------------------
# TAB 2: PDF-only Sales Assist
# --------------------------------------------------
with tab2:
    st.subheader("PDF + SKU Lookup → Sales Assist no container list")

    if not (sku_file and pdf_files):
        st.info("Upload SKU Lookup + PDFs to generate Sales Assist directly from PDFs.")
    else:
        if st.button("Parse PDFs + Match SKU + Build Sales Assist"):
            items_df = parse_pdfs_line_items(pdf_files, package_id_whitelist=None)

            if items_df.empty:
                st.error(
                    "No line-items were parsed from the PDFs.\n\n"
                    "If these are scanned images, OCR would be required.\n"
                    "If they are text-based, the line format may be different from the supported patterns."
                )
            else:
                sku_df = load_sku_lookup(sku_file)

                items_df["MAPPED DESCRIPTION"] = items_df["GRADE"].apply(map_description)

                items_df["MATCH KEY"] = (
                    items_df["MAPPED DESCRIPTION"] + "|"
                    + items_df["THICKNESS"].astype(str).apply(norm_int_str) + "|"
                    + items_df["WIDTH"].astype(str).apply(norm_int_str) + "|"
                    + items_df["LENGTH"].astype(str).apply(norm_int_str)
                )

                items_df = items_df.merge(
                    sku_df[["SKU", "MATCH KEY"]],
                    how="left",
                    on="MATCH KEY",
                )

                items_df["MATCH"] = items_df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")

                st.session_state.pdf_items_df = items_df
                st.session_state.pdf_sa_df = generate_sales_assist(items_df)

                st.success(f"Built Sales Assist from PDFs. Parsed {len(items_df)} line-items.")

        if st.session_state.pdf_items_df is not None:
            st.write("Parsed PDF line-items preview:")
            st.dataframe(st.session_state.pdf_items_df.head(200), use_container_width=True)

        if st.session_state.pdf_sa_df is not None:
            sa_name_pdf = st.text_input(
                "Sales Assist file name no extension",
                value="Sales_Assist_From_PDFs",
                key="sa_name_pdf",
            )

            st.download_button(
                "⬇️ Download Sales Assist Excel PDF-only",
                to_excel_bytes(st.session_state.pdf_sa_df),
                f"{sa_name_pdf}.xlsx",
            )


# --------------------------------------------------
# TAB 3: DABG LPN Assignment
# --------------------------------------------------
with tab3:
    st.subheader("DABG Package ID Assignment")

    st.write(
        "DABG now ignores grade for package assignment. "
        "It matches container rows to PDF LPN rows using `THICKNESS|WIDTH|LENGTH|PCS`."
    )

    st.write(
        "Example container key: `1|4|144|240`. "
        "Example consumed PDF key: `1|4|144|240|208`."
    )

    if not (container_file and sku_file and pdf_files):
        st.info("Upload Container List + SKU Lookup + PDFs to run DABG assignment.")
    else:
        if st.button("Run DABG LPN Assignment"):
            dabg_df, dabg_pdf_df = process_dabg(container_file, sku_file, pdf_files)

            st.session_state.dabg_df = dabg_df
            st.session_state.dabg_pdf_df = dabg_pdf_df
            st.session_state.dabg_sa_df = None

            assigned = int((dabg_df["DABG_PACKAGEID"].astype(str).str.strip() != "").sum())
            total = len(dabg_df)

            st.success(f"DABG assignment completed. Assigned {assigned} of {total} rows.")

        if st.session_state.dabg_pdf_df is not None:
            with st.expander("Parsed PDF LPN Pool"):
                if st.session_state.dabg_pdf_df.empty:
                    st.warning("No PDF LPN rows were parsed.")
                else:
                    st.dataframe(st.session_state.dabg_pdf_df, use_container_width=True)

                    st.download_button(
                        "⬇️ Download Parsed PDF LPN Pool",
                        to_excel_bytes(st.session_state.dabg_pdf_df),
                        "DABG_Parsed_PDF_LPN_Pool.xlsx",
                    )

        if st.session_state.dabg_df is not None:
            dabg_df = st.session_state.dabg_df

            try:
                st.dataframe(dabg_df.style.apply(highlight_dabg_mismatches, axis=1), use_container_width=True)
            except Exception:
                st.dataframe(dabg_df, use_container_width=True)

            assigned_count = int((dabg_df["DABG_PACKAGEID"].astype(str).str.strip() != "").sum())
            missing_lpn_count = int((dabg_df["DABG_PACKAGEID"].astype(str).str.strip() == "").sum())
            sku_missing_count = int((dabg_df["MATCH"].astype(str).str.upper() != "YES").sum())

            c1, c2, c3, c4 = st.columns(4)

            c1.metric("Rows", len(dabg_df))
            c2.metric("Assigned LPNs", assigned_count)
            c3.metric("Missing LPNs", missing_lpn_count)
            c4.metric("Missing SKU Matches", sku_missing_count)

            unmatched_lpn = dabg_df[dabg_df["DABG_PACKAGEID"].astype(str).str.strip() == ""]
            unmatched_sku = dabg_df[dabg_df["MATCH"].astype(str).str.upper() != "YES"]

            if len(unmatched_lpn) > 0:
                with st.expander("Rows with no available PDF LPN"):
                    st.dataframe(unmatched_lpn, use_container_width=True)

            if len(unmatched_sku) > 0:
                with st.expander("Rows with no SKU match"):
                    st.dataframe(unmatched_sku, use_container_width=True)

            st.download_button(
                "⬇️ Download DABG Match Excel",
                to_excel_bytes(dabg_df),
                container_file.name.replace(".xlsx", "_DABG_LPN_ASSIGNMENT.xlsx"),
            )

            st.divider()

            st.subheader("Sales Assist Export from DABG")

            st.write(
                "This export keeps the same Sales Assist columns, but uses consumed `DABG_PACKAGEID` "
                "values as the Sales Assist `Identifier`."
            )

            sa_name_dabg = st.text_input(
                "DABG Sales Assist file name no extension",
                value="Sales_Assist_DABG",
                key="sa_name_dabg",
            )

            usable_dabg = dabg_df[dabg_df["DABG_PACKAGEID"].astype(str).str.strip() != ""].copy()

            if len(usable_dabg) == 0:
                st.warning("No consumed DABG_PACKAGEID values are available for Sales Assist.")
            else:
                if st.button("Generate Sales Assist Excel DABG"):
                    st.session_state.dabg_sa_df = generate_sales_assist(
                        usable_dabg,
                        identifier_col="DABG_PACKAGEID",
                    )

                    st.success(f"Built DABG Sales Assist using {len(st.session_state.dabg_sa_df)} assigned LPN rows.")

                if st.session_state.dabg_sa_df is not None:
                    st.dataframe(st.session_state.dabg_sa_df, use_container_width=True)

                    st.download_button(
                        "⬇️ Download Sales Assist Excel DABG",
                        to_excel_bytes(st.session_state.dabg_sa_df),
                        f"{sa_name_dabg}.xlsx",
                    )
