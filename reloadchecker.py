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

if "page" not in st.session_state:
    st.session_state.page = "main"

if "dabg_df" not in st.session_state:
    st.session_state.dabg_df = None

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
    """
    Always return a Series aligned to df.index.
    Prevents .fillna errors when df.get(col, scalar) would return a scalar.
    """
    if df is None:
        return pd.Series([], dtype="object")

    if col in df.columns:
        return df[col]

    return pd.Series([default] * len(df), index=df.index)


# ==================================================
# SKU Logic
# ==================================================
def map_description(grade) -> str:
    """
    Maps container/PDF grade text to SKU lookup DESCRIPTION.

    Rules:
      - If grade blank -> DOG EAR
      - If contains APG -> TAEDA PINE APG
      - If contains DOG -> DOG EAR
      - If contains FENCE -> DOG EAR
      - If contains III or III/V or 3COM -> TAEDA PINE #3 COMMON
      - If contains D-Grade / GS-D-Grade / D GRADE -> TAEDA PINE D
      - Otherwise -> "" (don’t guess)
    """
    g = "" if pd.isna(grade) else str(grade).strip().upper()

    if g == "":
        return "DOG EAR"

    # D-grade mapping
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
            df["DESCRIPTION"] = df["DESCRIPTION"].str.upper().str.strip()

            df["MATCH KEY"] = (
                df["DESCRIPTION"] + "|"
                + df["THICKNESS"] + "|"
                + df["WIDTH"] + "|"
                + df["LENGTH"]
            )

            return df

    raise ValueError("SKU lookup missing required columns (SKU/DESCRIPTION/THICKNESS/WIDTH/LENGTH).")


# ==================================================
# PDF Header Extraction
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

    # order like 77660-2 / 81804-2 / 50050
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
# PDF Line Item Parser - Original Process
# ==================================================
def parse_pdfs_line_items(pdf_files, package_id_whitelist=None) -> pd.DataFrame:
    """
    Original parser used by the original workflow.

    Handles line formats like:
      GRADE DIM LPN PCS ...
      GRADE DIM PCS LPN ...

    DABG does NOT use this parser.
    DABG has its own parser below.
    """
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
                    thk = int(parts[0])
                    wid = int(parts[1])
                    leng = int(parts[2])
                except Exception:
                    continue

                # grade can be multi-token
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
# Full Process - Original
# ==================================================
def process_all(container_file, sku_file, pdf_files):
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    required_cols = {"PACKAGEID", "PCS", "GRADE", "THICKNESS", "WIDTH", "LENGTH"}
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise ValueError(f"Container list missing required columns: {missing}")

    # normalize container IDs
    df["PACKAGEID"] = df["PACKAGEID"].apply(norm_id)
    df["PCS"] = df["PCS"].apply(norm_int_str)

    package_whitelist = set(df["PACKAGEID"].astype(str))

    # Original process still uses original parser
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

    # SKU match
    sku_df = load_sku_lookup(sku_file)

    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)

    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|"
        + df["THICKNESS"].astype(str) + "|"
        + df["WIDTH"].astype(str) + "|"
        + df["LENGTH"].astype(str)
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


# ==================================================
# DABG Matching Logic
# ==================================================
def dabg_match_key_from_values(grade, thickness, width, length, pcs) -> str:
    """
    DABG match key:
      mapped grade/item | thickness | width | length | pcs

    This intentionally ignores container PACKAGEID and PDF LPN.
    """
    mapped_grade = map_description(grade)
    thk = norm_int_str(thickness)
    wid = norm_int_str(width)
    leng = norm_int_str(length)
    pieces = norm_int_str(pcs)

    return f"{mapped_grade}|{thk}|{wid}|{leng}|{pieces}"


def parse_dabg_pdf_lpns_by_item_dimension_pcs(pdf_files) -> pd.DataFrame:
    """
    DABG-only PDF parser.

    Built for warehouse receipt PDFs like:

        APG 1x8x144 (BUN: 120083 0.0000
        APG 1x4x144 (BUN: 240165 0.0000
        APG 1x6x144 (BUN: 160279 0.0000

    The combined BUN value is:
        PCS + LPN

    Examples:
        120083 -> PCS 120, LPN 083
        240165 -> PCS 240, LPN 165
        160279 -> PCS 160, LPN 279

    Returns:
        PACKAGEID = PDF LPN
        PCS
        QTY
        GRADE
        THICKNESS
        WIDTH
        LENGTH
        CONTAINER
        ORDERNUMBER
        PDF_FILE
    """
    rows = []

    # Match lines like:
    # APG 1x8x144 (BUN: 120083 0.0000
    bun_line_pat = re.compile(
        r"^\s*(.*?)\s+"
        r"(\d+)\s*[Xx]\s*(\d+)\s*[Xx]\s*(\d+)"
        r"\s+\(BUN:\s*([0-9]+)",
        re.IGNORECASE,
    )

    def split_pcs_lpn(combined: str):
        """
        For these DABG receipts, LPN is the last 3 digits.
        Everything before that is PCS.

        Examples:
            120083 -> PCS 120, LPN 083
            240165 -> PCS 240, LPN 165
            160279 -> PCS 160, LPN 279
        """
        combined = str(combined).strip()

        if not re.fullmatch(r"\d{4,}", combined):
            return "", ""

        lpn = combined[-3:]
        pcs = combined[:-3]

        if not pcs:
            return "", ""

        return lpn, pcs

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        pages_text = [(p.extract_text() or "") for p in reader.pages]
        full_text = "\n".join(pages_text)

        container, order = extract_container_and_order(full_text, pdf.name)

        for text in pages_text:
            for line in text.splitlines():
                line = line.strip()
                m = bun_line_pat.search(line)

                if not m:
                    continue

                grade = m.group(1).strip()
                thk = m.group(2).strip()
                wid = m.group(3).strip()
                leng = m.group(4).strip()
                combined = m.group(5).strip()

                lpn, pcs = split_pcs_lpn(combined)

                if not lpn or not pcs:
                    continue

                try:
                    pieces_int = int(pcs)
                    thk_int = int(thk)
                    wid_int = int(wid)
                    leng_int = int(leng)
                except Exception:
                    continue

                qty = int(round(pieces_int * (thk_int * wid_int * leng_int) / 144.0))

                rows.append(
                    {
                        "PACKAGEID": lpn,  # PDF LPN
                        "PCS": str(pieces_int),
                        "QTY": qty,
                        "GRADE": grade,
                        "THICKNESS": str(thk_int),
                        "WIDTH": str(wid_int),
                        "LENGTH": str(leng_int),
                        "CONTAINER": container,
                        "ORDERNUMBER": order,
                        "PDF_FILE": pdf.name,
                    }
                )

    return pd.DataFrame(rows)


def process_dabg(container_file, sku_file, pdf_files):
    """
    DABG workflow:
      - Reads container list
      - Reads warehouse receipt PDFs
      - Does NOT compare container PACKAGEID to PDF LPN
      - Matches by mapped grade/item, thickness, width, length, and PCS
      - Consumes/assigns one available PDF LPN per matching container row
      - Each PDF LPN can only be used once
    """
    raw_df = pd.read_excel(container_file, header=None, dtype=str)
    df = normalize_headers(raw_df).fillna("")

    required_cols = {"PACKAGEID", "PCS", "GRADE", "THICKNESS", "WIDTH", "LENGTH"}
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        raise ValueError(f"Container list missing required columns: {missing}")

    # Normalize container fields
    df["PACKAGEID"] = df["PACKAGEID"].apply(norm_id)
    df["PCS"] = df["PCS"].apply(norm_int_str)
    df["THICKNESS"] = df["THICKNESS"].apply(norm_int_str)
    df["WIDTH"] = df["WIDTH"].apply(norm_int_str)
    df["LENGTH"] = df["LENGTH"].apply(norm_int_str)

    # DABG uses its own parser.
    # Do NOT use parse_pdfs_line_items here.
    pdf_items = parse_dabg_pdf_lpns_by_item_dimension_pcs(pdf_files).fillna("")

    if pdf_items.empty:
        raise ValueError(
            "No DABG line-items were parsed from the PDFs. "
            "Expected lines like: APG 1x8x144 (BUN: 120083 0.0000"
        )

    # Keep PDF LPN leading zeroes.
    # Do NOT use norm_id on PDF PACKAGEID here.
    pdf_items["PACKAGEID"] = pdf_items["PACKAGEID"].astype(str).str.strip()
    pdf_items["PCS"] = pdf_items["PCS"].apply(norm_int_str)
    pdf_items["THICKNESS"] = pdf_items["THICKNESS"].apply(norm_int_str)
    pdf_items["WIDTH"] = pdf_items["WIDTH"].apply(norm_int_str)
    pdf_items["LENGTH"] = pdf_items["LENGTH"].apply(norm_int_str)

    # Build container-side DABG match key
    df["DABG MATCH KEY"] = df.apply(
        lambda r: dabg_match_key_from_values(
            r.get("GRADE", ""),
            r.get("THICKNESS", ""),
            r.get("WIDTH", ""),
            r.get("LENGTH", ""),
            r.get("PCS", ""),
        ),
        axis=1,
    )

    # Build PDF-side DABG match key
    pdf_items["DABG MATCH KEY"] = pdf_items.apply(
        lambda r: dabg_match_key_from_values(
            r.get("GRADE", ""),
            r.get("THICKNESS", ""),
            r.get("WIDTH", ""),
            r.get("LENGTH", ""),
            r.get("PCS", ""),
        ),
        axis=1,
    )

    # Create pools of available PDF LPNs by matching item/dim/pcs key.
    # This is the consume-available-LPNs logic.
    pdf_pools = {}

    for _, pdf_row in pdf_items.iterrows():
        key = pdf_row["DABG MATCH KEY"]

        if key not in pdf_pools:
            pdf_pools[key] = []

        pdf_pools[key].append(pdf_row.to_dict())

    assigned_lpns = []
    assigned_pdf_files = []
    assigned_pdf_containers = []
    assigned_pdf_grades = []
    assigned_pdf_thicknesses = []
    assigned_pdf_widths = []
    assigned_pdf_lengths = []
    assigned_pdf_pcs = []
    dabg_matches = []

    # Assign/consume one PDF LPN per matching container row
    for _, row in df.iterrows():
        key = row["DABG MATCH KEY"]
        available = pdf_pools.get(key, [])

        if available:
            pdf_match = available.pop(0)

            assigned_lpns.append(pdf_match.get("PACKAGEID", ""))
            assigned_pdf_files.append(pdf_match.get("PDF_FILE", ""))
            assigned_pdf_containers.append(pdf_match.get("CONTAINER", ""))
            assigned_pdf_grades.append(pdf_match.get("GRADE", ""))
            assigned_pdf_thicknesses.append(pdf_match.get("THICKNESS", ""))
            assigned_pdf_widths.append(pdf_match.get("WIDTH", ""))
            assigned_pdf_lengths.append(pdf_match.get("LENGTH", ""))
            assigned_pdf_pcs.append(pdf_match.get("PCS", ""))
            dabg_matches.append("YES")
        else:
            assigned_lpns.append("")
            assigned_pdf_files.append("")
            assigned_pdf_containers.append("")
            assigned_pdf_grades.append("")
            assigned_pdf_thicknesses.append("")
            assigned_pdf_widths.append("")
            assigned_pdf_lengths.append("")
            assigned_pdf_pcs.append("")
            dabg_matches.append("NO")

    df["DABG MATCH"] = dabg_matches

    # Keep original container PACKAGEID visible
    df["CONTAINER PACKAGEID"] = df["PACKAGEID"]

    # Assigned PDF LPN
    df["ASSIGNED PDF LPN"] = assigned_lpns
    df["ASSIGNED PDF FILE"] = assigned_pdf_files
    df["ASSIGNED PDF CONTAINER"] = assigned_pdf_containers

    # PDF-side audit values
    df["PDF GRADE"] = assigned_pdf_grades
    df["PDF THICKNESS"] = assigned_pdf_thicknesses
    df["PDF WIDTH"] = assigned_pdf_widths
    df["PDF LENGTH"] = assigned_pdf_lengths
    df["PDF PCS"] = assigned_pdf_pcs

    # Sales Assist identifier:
    # Use assigned PDF LPN when available.
    # Fallback to container PACKAGEID if no DABG match.
    df["DABG IDENTIFIER"] = df.apply(
        lambda r: r["ASSIGNED PDF LPN"]
        if str(r.get("ASSIGNED PDF LPN", "")).strip()
        else r.get("PACKAGEID", ""),
        axis=1,
    )

    # SKU match - same as original process
    sku_df = load_sku_lookup(sku_file)

    df["MAPPED DESCRIPTION"] = df["GRADE"].apply(map_description)

    df["MATCH KEY"] = (
        df["MAPPED DESCRIPTION"] + "|"
        + df["THICKNESS"].astype(str) + "|"
        + df["WIDTH"].astype(str) + "|"
        + df["LENGTH"].astype(str)
    )

    df = df.merge(sku_df[["SKU", "MATCH KEY"]], how="left", on="MATCH KEY")
    df["MATCH"] = df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")
    df = df.fillna("")

    df["DABG MATCHED BY"] = "ITEM/DIMENSION/PCS"

    dabg_cols = [
        "CONTAINER PACKAGEID",
        "ASSIGNED PDF LPN",
        "DABG MATCH",
        "DABG MATCHED BY",
        "ASSIGNED PDF FILE",
        "ASSIGNED PDF CONTAINER",
        "PDF GRADE",
        "PDF THICKNESS",
        "PDF WIDTH",
        "PDF LENGTH",
        "PDF PCS",
        "DABG IDENTIFIER",
        "SKU",
        "MATCH",
    ]

    existing_dabg = [c for c in dabg_cols if c in df.columns]
    others = [c for c in df.columns if c not in existing_dabg]

    return df[others + existing_dabg]


# ==================================================
# PCS Fix
# ==================================================
def fix_pcs_mismatch_use_container_truth(df: pd.DataFrame):
    """
    Container list PCS is truth.
    Fix ONLY the audit columns:
      PCS CHECK <- PCS where PCS CHECK exists and differs
      PCS MATCH <- YES

    Does NOT modify PCS.
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
# Sales Assist Generator - Original
# ==================================================
def generate_sales_assist(df: pd.DataFrame) -> pd.DataFrame:
    # OrderNumber
    order_raw = col_or_default(df, "ORDERNUMBER", "").astype(str).str.split("-").str[0].str.strip()

    if series_digits_only(order_raw):
        order_out = pd.to_numeric(order_raw, errors="coerce").fillna(0).astype(int)
    else:
        order_out = order_raw

    # Identifier from original PACKAGEID
    ident_raw = col_or_default(df, "PACKAGEID", "").astype(str).str.strip()

    if series_digits_only(ident_raw):
        ident_out = pd.to_numeric(ident_raw, errors="coerce").fillna(0).astype(int)
    else:
        ident_out = ident_raw

    # Pieces
    pcs = pd.to_numeric(col_or_default(df, "PCS", 0), errors="coerce").fillna(0).astype(int)

    # Quantity BF
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
# Sales Assist Generator - DABG
# ==================================================
def generate_dabg_sales_assist(df: pd.DataFrame) -> pd.DataFrame:
    # OrderNumber
    order_raw = col_or_default(df, "ORDERNUMBER", "").astype(str).str.split("-").str[0].str.strip()

    if series_digits_only(order_raw):
        order_out = pd.to_numeric(order_raw, errors="coerce").fillna(0).astype(int)
    else:
        order_out = order_raw

    # Identifier should be assigned PDF LPN for DABG
    ident_raw = col_or_default(df, "DABG IDENTIFIER", "").astype(str).str.strip()

    # IMPORTANT:
    # Do not convert DABG identifiers to int, because PDF LPNs can have leading zeroes.
    ident_out = ident_raw

    # Pieces
    pcs = pd.to_numeric(col_or_default(df, "PCS", 0), errors="coerce").fillna(0).astype(int)

    # Quantity BF
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

            # DABG audit visibility
            "ContainerList_PACKAGEID": col_or_default(df, "CONTAINER PACKAGEID", ""),
            "Assigned_PDF_LPN": col_or_default(df, "ASSIGNED PDF LPN", ""),
            "DABG_MATCH": col_or_default(df, "DABG MATCH", ""),
            "Assigned_PDF_Container": col_or_default(df, "ASSIGNED PDF CONTAINER", ""),
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
    if row.get("DABG MATCH") != "YES" or row.get("MATCH") != "YES":
        return ["background-color: #ffcccc"] * len(row)

    return [""] * len(row)


# ==================================================
# Streamlit UI
# ==================================================
st.set_page_config(page_title="BIFP Import Checker", layout="wide")
st.title("📦 BIFP SKU + Receive + PCS Match + Sales Assist")

nav1, nav2 = st.columns([1, 1])

with nav1:
    if st.button("Original Process"):
        st.session_state.page = "main"

with nav2:
    if st.button("DABG"):
        st.session_state.page = "dabg"


# ==================================================
# ORIGINAL PROCESS PAGE
# ==================================================
if st.session_state.page == "main":
    container_file = st.file_uploader(
        "Upload Container List Excel (optional)",
        type="xlsx",
        key="main_container_file",
    )

    sku_file = st.file_uploader(
        "Upload SKU Lookup Excel",
        type="xlsx",
        key="main_sku_file",
    )

    pdf_files = st.file_uploader(
        "Upload PDF Files",
        type="pdf",
        accept_multiple_files=True,
        key="main_pdf_files",
    )

    tab1, tab2 = st.tabs(
        [
            "Full Match + Audit (Container + PDFs)",
            "PDF + SKU Lookup → Sales Assist (no container)",
        ]
    )

    # --------------------------------------------------
    # TAB 1: Full process
    # --------------------------------------------------
    with tab1:
        st.subheader("Full Match + Audit")

        if not (container_file and sku_file and pdf_files):
            st.info("Upload **Container List + SKU Lookup + PDFs** to run the full match.")
        else:
            if st.button("Run Full Process"):
                try:
                    st.session_state.processed_df = process_all(container_file, sku_file, pdf_files)
                    st.success("Full process completed")
                except Exception as e:
                    st.error(str(e))

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
                            st.info("No PCS mismatches found to fix or PCS CHECK was blank.")
                        else:
                            st.success(f"Fixed {n} PCS mismatches using **container PCS as truth**.")

                st.divider()

                st.subheader("Sales Assist Export (from Full Match)")

                sa_name = st.text_input(
                    "Sales Assist file name (no extension)",
                    value="Sales_Assist_Full",
                    key="main_sa_name",
                )

                if st.button("Generate Sales Assist Excel (Full Match)"):
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
        st.subheader("PDF + SKU Lookup → Sales Assist (no container list)")

        if not (sku_file and pdf_files):
            st.info("Upload **SKU Lookup + PDFs** to generate Sales Assist directly from PDFs.")
        else:
            if st.button("Parse PDFs + Match SKU + Build Sales Assist"):
                try:
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
                            + items_df["THICKNESS"].astype(str) + "|"
                            + items_df["WIDTH"].astype(str) + "|"
                            + items_df["LENGTH"].astype(str)
                        )

                        items_df = items_df.merge(
                            sku_df[["SKU", "MATCH KEY"]],
                            how="left",
                            on="MATCH KEY",
                        )

                        items_df["MATCH"] = items_df["SKU"].apply(lambda x: "YES" if sku_is_valid(x) else "NO")

                        st.session_state.pdf_items_df = items_df
                        st.session_state.pdf_sa_df = generate_sales_assist(items_df)

                        st.success(f"Built Sales Assist from PDFs ({len(items_df)} line-items).")

                except Exception as e:
                    st.error(str(e))

            if st.session_state.pdf_items_df is not None:
                st.write("Parsed PDF line-items preview:")
                st.dataframe(st.session_state.pdf_items_df.head(200), use_container_width=True)

            if st.session_state.pdf_sa_df is not None:
                sa_name_pdf = st.text_input(
                    "Sales Assist file name (no extension)",
                    value="Sales_Assist_From_PDFs",
                    key="pdf_sa_name",
                )

                st.download_button(
                    "⬇️ Download Sales Assist Excel (PDF-only)",
                    to_excel_bytes(st.session_state.pdf_sa_df),
                    f"{sa_name_pdf}.xlsx",
                )


# ==================================================
# DABG PAGE
# ==================================================
elif st.session_state.page == "dabg":
    st.subheader("DABG Matching Workflow")

    st.info(
        "DABG keeps the original process separate. "
        "It does not match container PACKAGEID to PDF LPN. "
        "It matches by mapped grade/item, thickness, width, length, and PCS, "
        "then assigns the first available PDF LPN from the imported PDFs."
    )

    dabg_container_file = st.file_uploader(
        "Upload Container List Excel",
        type="xlsx",
        key="dabg_container_file",
    )

    dabg_sku_file = st.file_uploader(
        "Upload SKU Lookup Excel",
        type="xlsx",
        key="dabg_sku_file",
    )

    dabg_pdf_files = st.file_uploader(
        "Upload PDF Files",
        type="pdf",
        accept_multiple_files=True,
        key="dabg_pdf_files",
    )

    if not (dabg_container_file and dabg_sku_file and dabg_pdf_files):
        st.info("Upload **Container List + SKU Lookup + PDFs** to run DABG.")
    else:
        if st.button("Run DABG Match"):
            try:
                st.session_state.dabg_df = process_dabg(
                    dabg_container_file,
                    dabg_sku_file,
                    dabg_pdf_files,
                )
                st.session_state.dabg_sa_df = None
                st.success("DABG match completed")
            except Exception as e:
                st.error(str(e))

        if st.session_state.dabg_df is not None:
            dabg_df = st.session_state.dabg_df

            try:
                st.dataframe(
                    dabg_df.style.apply(highlight_dabg_mismatches, axis=1),
                    use_container_width=True,
                )
            except Exception:
                st.dataframe(dabg_df, use_container_width=True)

            st.download_button(
                "⬇️ Download DABG Match Excel",
                to_excel_bytes(dabg_df),
                dabg_container_file.name.replace(".xlsx", "_DABG_MATCH.xlsx"),
            )

            st.divider()

            st.subheader("Sales Assist Export - DABG")

            dabg_sa_name = st.text_input(
                "DABG Sales Assist file name (no extension)",
                value="Sales_Assist_DABG",
                key="dabg_sa_name",
            )

            if st.button("Generate DABG Sales Assist Excel"):
                st.session_state.dabg_sa_df = generate_dabg_sales_assist(dabg_df)

            if st.session_state.dabg_sa_df is not None:
                st.download_button(
                    "⬇️ Download DABG Sales Assist Excel",
                    to_excel_bytes(st.session_state.dabg_sa_df),
                    f"{dabg_sa_name}.xlsx",
                )
