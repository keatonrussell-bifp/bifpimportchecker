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

if "dabg_raw_pdf_lines_df" not in st.session_state:
    st.session_state.dabg_raw_pdf_lines_df = None


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


def clean_pdf_line(line: str) -> str:
    return " ".join(str(line).replace("\xa0", " ").strip().split())


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
# Receive Summary Ticket Parser (Grane Laramie)
# ==================================================
def is_receive_summary_ticket(full_text: str) -> bool:
    """Identify the new Grane Laramie receive-summary layout."""
    up = (full_text or "").upper()
    return "RECEIVE SUMMARY TICKET" in up and "DETAILS:" in up and "LOT#:" in up


def parse_receive_summary_ticket(
    pdf_name: str,
    pages_text,
    package_id_whitelist=None,
):
    """
    Parse Birmingham International / Grane Laramie Receive Summary Tickets.

    The item dimensions and grade are printed once, followed by one or more
    detail rows. Each detail row becomes one normal import-generator row:

      PACKAGEID = Lot#
      PCS       = detail Qty
      QTY       = board feet calculated from PCS and dimensions

    Item context is kept across page breaks because some receipts continue a
    detail group on the following page before printing the next item heading.
    """
    dim_pat = re.compile(
        r'^\s*'
        r'(?P<thk>\d+(?:[\.,]\d+)?)\s*["”]\s*'
        r'(?P<wid>\d+(?:[\.,]\d+)?)\s*["”]\s*'
        r'(?P<len>\d+(?:[\.,]\d+)?)\s*["”]\s*$'
    )
    lot_pat = re.compile(r'\bLOT\s*#?\s*:\s*([A-Za-z0-9\-]+)', re.IGNORECASE)
    qty_pat = re.compile(r'\bQTY\s*:\s*([\d,]+)', re.IGNORECASE)

    def dim_to_int(value: str) -> str:
        value = str(value).strip().replace(',', '.')
        try:
            return str(int(round(float(value))))
        except Exception:
            return ""

    def grade_from_description(text: str) -> str:
        up = " ".join(str(text).upper().split())

        if re.search(r'\bAPG\b', up):
            return "APG"

        if re.search(r'\b3\s*COM\b|\b3COM\b|\bIII/V\b|\bIII\b', up):
            return "3COM"

        if re.search(r'\bD[-\s]?GRADE\b|\bGRADE[-\s]?D\b', up):
            return "D-GRADE"

        if re.search(r'\bFB\b|\bDOG\b|\bFENCE\b', up):
            return "FB"

        return up.strip()

    # Filename is the safest source when the trailer field is blank. If the
    # filename is not a container number, fall back to the document text.
    filename_match = re.search(r'\b([A-Z]{4}\d{7})\b', (pdf_name or '').upper())
    full_text = "\n".join(pages_text)
    text_match = re.search(r'\b([A-Z]{4}\d{7})\b', full_text.upper())
    container = (
        filename_match.group(1)
        if filename_match
        else (text_match.group(1) if text_match else "")
    )

    # These tickets show the reload/vendor PO, not the internal Sales Assist
    # order number. In the full-match workflow, ORDERNUMBER comes from the
    # uploaded container Excel, so leave this blank rather than using a Lot#
    # or the vendor PO as an incorrect order number.
    order = ""

    lines = []
    for page_num, text in enumerate(pages_text, start=1):
        for line_num, line in enumerate((text or "").splitlines(), start=1):
            cleaned = clean_pdf_line(line)
            if cleaned:
                lines.append((page_num, line_num, cleaned))

    rows = []
    current_item = None
    i = 0

    while i < len(lines):
        page_num, line_num, line = lines[i]
        dim_match = dim_pat.match(line)

        if dim_match:
            thk = dim_to_int(dim_match.group('thk'))
            wid = dim_to_int(dim_match.group('wid'))
            leng = dim_to_int(dim_match.group('len'))

            description_parts = []
            j = i + 1

            # The item description may wrap across several extracted lines.
            # Stop before detail rows or the next item heading.
            while j < len(lines) and j <= i + 8:
                _, _, nxt = lines[j]
                nxt_up = nxt.upper().strip()

                if dim_pat.match(nxt) or lot_pat.search(nxt):
                    break

                if nxt_up in {"PALLET", "TOTALS :", "TOTALS:"}:
                    break

                # Ignore table values and generic header fragments.
                if re.fullmatch(r'[\d,]+', nxt):
                    j += 1
                    continue

                if nxt_up in {
                    "SKU", "QUALIFIER", "ITEM DESCRIPTION", "INVENTORY",
                    "INVENTORY QTY", "VARIABLE QTY", "CU FT", "LBS",
                    "PACKED PER DIM UOM", "TOTAL QTY DIM UOM",
                    "DIM UNIT OF MEASURE", "BLI",
                }:
                    j += 1
                    continue

                description_parts.append(nxt)
                j += 1

            grade = grade_from_description(" ".join(description_parts))

            current_item = {
                "GRADE": grade,
                "THICKNESS": thk,
                "WIDTH": wid,
                "LENGTH": leng,
            }

            i += 1
            continue

        lot_match = lot_pat.search(line)

        if lot_match and current_item is not None:
            package_id = norm_id(lot_match.group(1))
            pieces = ""

            qty_match = qty_pat.search(line)
            if qty_match:
                pieces = norm_int_str(qty_match.group(1))

            # PyPDF2 normally puts Qty on the following line. Scan only a few
            # lines and stop if a new lot/item starts.
            if not pieces:
                for j in range(i + 1, min(i + 5, len(lines))):
                    _, _, nxt = lines[j]

                    if dim_pat.match(nxt) or lot_pat.search(nxt):
                        break

                    qty_match = qty_pat.search(nxt)
                    if qty_match:
                        pieces = norm_int_str(qty_match.group(1))
                        break

            if package_id and pieces:
                thk = int(current_item["THICKNESS"])
                wid = int(current_item["WIDTH"])
                leng = int(current_item["LENGTH"])
                pcs_int = int(pieces)
                qty = int(round(pcs_int * (thk * wid * leng) / 144.0))

                rows.append(
                    {
                        "PACKAGEID": package_id,
                        "PCS": pcs_int,
                        "QTY": qty,
                        "GRADE": current_item["GRADE"],
                        "THICKNESS": str(thk),
                        "WIDTH": str(wid),
                        "LENGTH": str(leng),
                        "CONTAINER": container,
                        "ORDERNUMBER": order,
                        "PDF_FILE": pdf_name,
                    }
                )

        i += 1

    if not rows:
        return []

    # One physical lot should appear only once in a receipt. This also protects
    # against PDF extraction occasionally repeating a line at a page boundary.
    deduped = []
    seen = set()

    for row in rows:
        key = (row["PDF_FILE"], row["PACKAGEID"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(row)

    return deduped


# ==================================================
# Existing PDF Line Item Parser + New Format Dispatch
# ==================================================
def parse_pdfs_line_items(pdf_files, package_id_whitelist=None) -> pd.DataFrame:
    """
    Parse all supported receipt formats.

    Existing receipt parsing remains unchanged. The only addition is a format
    check that routes Grane Laramie Receive Summary Tickets to their dedicated
    parser, then returns the same output columns used by the existing app.
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

        if is_receive_summary_ticket(full_text):
            rows.extend(
                parse_receive_summary_ticket(
                    pdf_name=pdf.name,
                    pages_text=pages_text,
                    package_id_whitelist=package_id_whitelist,
                )
            )
            continue

        # Legacy path below is intentionally unchanged.
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


def extract_raw_pdf_lines(pdf_files) -> pd.DataFrame:
    rows = []

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))

        for page_num, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""

            for line_num, line in enumerate(text.splitlines(), start=1):
                line_clean = clean_pdf_line(line)

                if line_clean:
                    rows.append(
                        {
                            "PDF_FILE": pdf.name,
                            "PDF_PAGE": page_num,
                            "PDF_LINE": line_num,
                            "TEXT": line_clean,
                        }
                    )

    return pd.DataFrame(rows)

def parse_dabg_pdfs_lpn_rows(pdf_files, valid_container_keys=None):
    """
    Parses warehouse receipt PDFs into LPN rows for DABG.

    DABG key ignores grade:
      DABG_MATCH_KEY = THICKNESS|WIDTH|LENGTH|PCS

    On these North Florida Warehouse receipts the LPN is a two-part value
    such as "CHS 5772" (an alpha prefix plus a number). PyPDF2 emits the
    prefix and the number as SEPARATE tokens, and the column order varies
    between files, e.g.:

        APG 1x6x144 (BUN: 160 CHS 6964 0.0000   ->  PIECES=160  LPN="CHS 6964"
        APG 1x6x144 (BUN: CHS 5772 160 0.0000   ->  LPN="CHS 5772"  PIECES=160

    resolve_dabg_lpn_pcs rebuilds the FULL LPN (prefix + number) and uses the
    container dim/pcs key list to pin down which number is the piece count.
    """
    rows = []
    raw_lines_rows = []

    if valid_container_keys is None:
        valid_container_keys = set()
    else:
        valid_container_keys = set(str(x).strip() for x in valid_container_keys if str(x).strip())

    item_pat = re.compile(
        r"""
        (?P<grade>[A-Za-z0-9#/\-\s]+?)
        \s+
        (?P<thickness>\d+)
        \s*[xX]\s*
        (?P<width>\d+)
        \s*[xX]\s*
        (?P<length>\d+)
        """,
        re.VERBOSE,
    )

    header_words = {
        "ITEM", "LOT", "LOT CODE", "SUBLOT", "SUBLOT CODE",
        "LPN", "PIECES", "TOTAL", "TOTAL LBS", "WAREHOUSE",
        "RECEIPT", "PAGE", "TRANSACTION", "CONTAINER", "CARRIER",
        "RECVD", "FROM", "FOR", "ACCOUNT", "SPECIAL", "INSTRUCTIONS",
        "BUN", "BUNDLE",
    }

    def looks_like_header(line: str) -> bool:
        up = line.upper().strip()

        if up in header_words:
            return True

        if "ITEM" in up and "LPN" in up and "PIECES" in up:
            return True

        if "WAREHOUSE RECEIPT" in up:
            return True

        if up.startswith("PAGE "):
            return True

        if up.startswith("TRANSACTION"):
            return True

        if up.startswith("RECEIPT"):
            return True

        if up.startswith("CONTAINER"):
            return True

        if up.startswith("P.O."):
            return True

        return False

    def token_is_total_lbs(tok: str) -> bool:
        tok = str(tok).strip()
        return bool(re.fullmatch(r"\d+\.\d+", tok))

    def numeric_or_alnum_tokens(text: str):
        return re.findall(r"[A-Za-z0-9\-]+|\d+\.\d+", text)

    def is_valid_pcs(tok: str) -> bool:
        tok = str(tok).strip()

        if not re.fullmatch(r"\d+", tok):
            return False

        n = int(tok)

        return 1 <= n <= 10000

    def is_valid_lpn(tok: str) -> bool:
        tok = str(tok).strip()

        if not tok:
            return False

        if tok.upper() in header_words:
            return False

        if token_is_total_lbs(tok):
            return False

        # Allow an internal space so two-part LPNs like "CHS 5772" survive.
        return bool(re.fullmatch(r"[A-Za-z0-9\- ]+", tok))

    def choose_lpn_and_pcs(thk, wid, leng, token_a, token_b):
        """
        Legacy two-token resolver. Kept for reference; the DABG path now uses
        resolve_dabg_lpn_pcs so the full LPN (prefix + number) is preserved.
        """
        a = norm_id(token_a)
        b = norm_id(token_b)

        key_if_b_is_pcs = make_dabg_dim_pcs_key(thk, wid, leng, b)
        key_if_a_is_pcs = make_dabg_dim_pcs_key(thk, wid, leng, a)

        b_matches_container = key_if_b_is_pcs in valid_container_keys
        a_matches_container = key_if_a_is_pcs in valid_container_keys

        if b_matches_container and not a_matches_container:
            return a, b, "container-key-selected-second-token-as-pcs"

        if a_matches_container and not b_matches_container:
            return b, a, "container-key-selected-first-token-as-pcs"

        if b_matches_container and a_matches_container:
            return a, b, "both-possible-default-second-token-as-pcs"

        return a, b, "no-container-key-match-default-second-token-as-pcs"

    def resolve_dabg_lpn_pcs(thk, wid, leng, toks):
        """
        Rebuild the FULL LPN and identify PCS from the post-dimension tokens.

        - prefix      = first pure-alpha token (e.g. "CHS")
        - lpn_number  = the digits that belong with the prefix
        - pcs         = the numeric token whose THK|WID|LEN|PCS key is in the
                        container list; this is what disambiguates pcs vs lpn#

        The full LPN is returned as "<prefix> <lpn_number>", e.g. "CHS 5772".
        """
        toks = [str(t).strip() for t in toks if str(t).strip()]

        prefix = ""
        prefix_idx = None
        for idx, t in enumerate(toks):
            if re.fullmatch(r"[A-Za-z]+", t):
                prefix = t
                prefix_idx = idx
                break

        numerics = [t for t in toks if re.fullmatch(r"\d+", t)]

        # PCS = the numeric whose dimension+pcs key exists in the container list.
        pcs = ""
        for n in numerics:
            if make_dabg_dim_pcs_key(thk, wid, leng, n) in valid_container_keys:
                pcs = n
                break

        # LPN number = the digits paired with the prefix.
        # Prefer the first numeric AFTER the prefix that is not the PCS,
        # then the first numeric BEFORE the prefix that is not the PCS,
        # then any remaining numeric.
        lpn_number = ""
        if prefix_idx is not None:
            for t in toks[prefix_idx + 1:]:
                if re.fullmatch(r"\d+", t) and t != pcs:
                    lpn_number = t
                    break
            if not lpn_number:
                for t in reversed(toks[:prefix_idx]):
                    if re.fullmatch(r"\d+", t) and t != pcs:
                        lpn_number = t
                        break
        if not lpn_number:
            for t in numerics:
                if t != pcs:
                    lpn_number = t
                    break

        if prefix and lpn_number:
            lpn = prefix + " " + lpn_number
            decision = "dabg-prefix-merged-lpn"
        elif prefix:
            lpn = prefix
            decision = "dabg-prefix-only-no-number"
        else:
            lpn = lpn_number
            decision = "dabg-numeric-lpn"

        # Fallback PCS when the container list has no matching key.
        if not pcs:
            leftover = [t for t in numerics if t != lpn_number]
            if leftover:
                pcs = leftover[0]
                decision = decision + "-pcs-fallback"

        return lpn, pcs, prefix, lpn_number, decision

    def add_row(pdf_name, page_num, line_num, grade_raw, thk, wid, leng, toks, container, order, source_method):
        thk = norm_int_str(thk)
        wid = norm_int_str(wid)
        leng = norm_int_str(leng)

        if not thk or not wid or not leng:
            return

        lpn, pcs, prefix, lpn_number, decision = resolve_dabg_lpn_pcs(thk, wid, leng, toks)

        pcs = norm_int_str(pcs)
        lpn = norm_id(lpn)

        if not pcs or not lpn:
            return

        if not is_valid_pcs(pcs):
            return

        if not is_valid_lpn(lpn):
            return

        key = make_dabg_dim_pcs_key(thk, wid, leng, pcs)

        rows.append(
            {
                "PDF_FILE": pdf_name,
                "PDF_PAGE": page_num,
                "PDF_LINE": line_num,
                "PDF_GRADE_RAW": grade_raw,
                "THICKNESS": thk,
                "WIDTH": wid,
                "LENGTH": leng,
                "PCS": pcs,
                "LPN": lpn,
                "DABG_MATCH_KEY": key,
                "DABG_MATCH_KEY_LPN": f"{key}|{lpn}",
                "CONTAINER": container,
                "ORDERNUMBER": order,
                "SOURCE_METHOD": source_method,
                "TOKEN_A": prefix,
                "TOKEN_B": lpn_number,
                "PARSE_DECISION": decision,
            }
        )

    for pdf in pdf_files:
        reader = PdfReader(BytesIO(pdf.getvalue()))
        pages_text = [(p.extract_text() or "") for p in reader.pages]
        full_text = "\n".join(pages_text)

        container, order = extract_container_and_order(full_text, pdf.name)

        for page_num, text in enumerate(pages_text, start=1):
            lines = [clean_pdf_line(x) for x in text.splitlines()]
            lines = [x for x in lines if x]

            for line_num, line in enumerate(lines, start=1):
                raw_lines_rows.append(
                    {
                        "PDF_FILE": pdf.name,
                        "PDF_PAGE": page_num,
                        "PDF_LINE": line_num,
                        "TEXT": line,
                    }
                )

            i = 0

            while i < len(lines):
                line = lines[i]

                if looks_like_header(line):
                    i += 1
                    continue

                m = item_pat.search(line)

                if not m:
                    i += 1
                    continue

                grade_raw = m.group("grade").strip()
                thk = m.group("thickness")
                wid = m.group("width")
                leng = m.group("length")

                after = line[m.end():].strip()
                after = after.replace("(", " ").replace(")", " ").replace(":", " ")

                toks = numeric_or_alnum_tokens(after)

                toks = [
                    t for t in toks
                    if t.upper() not in {"BUN", "BUNDLE", "LPN", "PIECES", "TOTAL", "LBS"}
                ]

                toks = [
                    t for t in toks
                    if not token_is_total_lbs(t)
                ]

                # Same-line case.
                # The tokens after the dimension contain, in some order:
                #   PIECES, the LPN prefix (CHS), and the LPN number.
                # resolve_dabg_lpn_pcs figures out which is which and rebuilds
                # the full LPN, so we hand it the whole token list.
                if len(toks) >= 2:
                    add_row(
                        pdf.name,
                        page_num,
                        i + 1,
                        grade_raw,
                        thk,
                        wid,
                        leng,
                        toks,
                        container,
                        order,
                        "same-line-prefix-merged",
                    )

                    i += 1
                    continue

                # Stacked extraction case:
                # APG 1x8x144 (BUN:
                # 160
                # CHS
                # 5772
                # 0.0000
                stacked_tokens = []

                j = i + 1

                while j < len(lines) and j <= i + 8:
                    nxt = lines[j]

                    if item_pat.search(nxt):
                        break

                    if not looks_like_header(nxt):
                        for tok in numeric_or_alnum_tokens(nxt):
                            if tok.upper() in {"BUN", "BUNDLE", "LPN", "PIECES", "TOTAL", "LBS"}:
                                continue

                            if token_is_total_lbs(tok):
                                continue

                            stacked_tokens.append(tok)

                    # Collect enough tokens (pieces + prefix + lpn number) before stopping.
                    if len(stacked_tokens) >= 4:
                        break

                    j += 1

                if len(stacked_tokens) >= 2:
                    add_row(
                        pdf.name,
                        page_num,
                        i + 1,
                        grade_raw,
                        thk,
                        wid,
                        leng,
                        stacked_tokens,
                        container,
                        order,
                        "stacked-lines-prefix-merged",
                    )

                    i = j + 1
                    continue

                i += 1

    pdf_df = pd.DataFrame(rows)
    raw_lines_df = pd.DataFrame(raw_lines_rows)

    if not pdf_df.empty:
        pdf_df = pdf_df.drop_duplicates(
            subset=["PDF_FILE", "PDF_PAGE", "PDF_LINE", "LPN", "PCS", "DABG_MATCH_KEY"],
            keep="first",
        ).reset_index(drop=True)

    return pdf_df, raw_lines_df
 


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

    df["DABG_CONTAINER_MATCH_KEY"] = df.apply(
        lambda r: make_dabg_dim_pcs_key(
            r.get("THICKNESS", ""),
            r.get("WIDTH", ""),
            r.get("LENGTH", ""),
            r.get("PCS", ""),
        ),
        axis=1,
    )

    valid_container_keys = set(df["DABG_CONTAINER_MATCH_KEY"].astype(str).str.strip())

    pdf_df, raw_pdf_lines_df = parse_dabg_pdfs_lpn_rows(
        pdf_files,
        valid_container_keys=valid_container_keys,
    )

    lpn_pool = defaultdict(deque)

    if not pdf_df.empty:
        pdf_df = pdf_df.reset_index(drop=True)

        for _, r in pdf_df.iterrows():
            key = r["DABG_MATCH_KEY"]
            lpn = r["LPN"]
            lpn_pool[key].append(lpn)

    assigned_lpns = []
    source_keys = []

    for _, r in df.iterrows():
        key = r["DABG_CONTAINER_MATCH_KEY"]

        if lpn_pool[key]:
            lpn = lpn_pool[key].popleft()
            assigned_lpns.append(lpn)
            source_keys.append(f"{key}|{lpn}")
        else:
            assigned_lpns.append("")
            source_keys.append("")

    df["DABG_MATCH_KEY"] = df["DABG_CONTAINER_MATCH_KEY"]
    df["DABG_PACKAGEID"] = assigned_lpns
    df["DABG_MATCH_KEY_LPN"] = source_keys
    df["DABG LPN MATCH"] = df["DABG_PACKAGEID"].apply(lambda x: "YES" if str(x).strip() else "NO")

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

    return df, pdf_df, raw_pdf_lines_df


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
        "DABG ignores grade for package assignment. "
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
            dabg_df, dabg_pdf_df, raw_pdf_lines_df = process_dabg(container_file, sku_file, pdf_files)

            st.session_state.dabg_df = dabg_df
            st.session_state.dabg_pdf_df = dabg_pdf_df
            st.session_state.dabg_raw_pdf_lines_df = raw_pdf_lines_df
            st.session_state.dabg_sa_df = None

            assigned = int((dabg_df["DABG_PACKAGEID"].astype(str).str.strip() != "").sum())
            total = len(dabg_df)
            parsed_pdf_rows = 0 if dabg_pdf_df is None or dabg_pdf_df.empty else len(dabg_pdf_df)

            st.success(
                f"DABG assignment completed. Assigned {assigned} of {total} rows. "
                f"Parsed {parsed_pdf_rows} PDF LPN rows."
            )

        if st.session_state.dabg_pdf_df is not None:
            with st.expander("Parsed PDF LPN Pool", expanded=True):
                if st.session_state.dabg_pdf_df.empty:
                    st.warning(
                        "No PDF LPN rows were parsed. Open Raw PDF Extracted Lines below and check how PyPDF2 is reading the warehouse receipt."
                    )
                else:
                    st.dataframe(st.session_state.dabg_pdf_df, use_container_width=True)

                    st.download_button(
                        "⬇️ Download Parsed PDF LPN Pool",
                        to_excel_bytes(st.session_state.dabg_pdf_df),
                        "DABG_Parsed_PDF_LPN_Pool.xlsx",
                    )

                    key_counts = (
                        st.session_state.dabg_pdf_df
                        .groupby("DABG_MATCH_KEY", as_index=False)
                        .agg(PDF_LPN_COUNT=("LPN", "count"))
                        .sort_values("DABG_MATCH_KEY")
                    )

                    st.write("PDF LPN counts by DABG key:")
                    st.dataframe(key_counts, use_container_width=True)

        if st.session_state.dabg_raw_pdf_lines_df is not None:
            with st.expander("Raw PDF Extracted Lines"):
                st.dataframe(st.session_state.dabg_raw_pdf_lines_df, use_container_width=True)

                st.download_button(
                    "⬇️ Download Raw PDF Extracted Lines",
                    to_excel_bytes(st.session_state.dabg_raw_pdf_lines_df),
                    "DABG_Raw_PDF_Extracted_Lines.xlsx",
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

            if st.session_state.dabg_pdf_df is not None and not st.session_state.dabg_pdf_df.empty:
                container_keys = (
                    dabg_df
                    .groupby("DABG_CONTAINER_MATCH_KEY", as_index=False)
                    .agg(CONTAINER_ROW_COUNT=("DABG_CONTAINER_MATCH_KEY", "count"))
                    .rename(columns={"DABG_CONTAINER_MATCH_KEY": "DABG_MATCH_KEY"})
                )

                pdf_keys = (
                    st.session_state.dabg_pdf_df
                    .groupby("DABG_MATCH_KEY", as_index=False)
                    .agg(PDF_LPN_COUNT=("LPN", "count"))
                )

                key_compare = container_keys.merge(pdf_keys, how="outer", on="DABG_MATCH_KEY").fillna(0)
                key_compare["CONTAINER_ROW_COUNT"] = key_compare["CONTAINER_ROW_COUNT"].astype(int)
                key_compare["PDF_LPN_COUNT"] = key_compare["PDF_LPN_COUNT"].astype(int)
                key_compare["SHORTAGE"] = key_compare["CONTAINER_ROW_COUNT"] - key_compare["PDF_LPN_COUNT"]

                with st.expander("DABG Key Compare Container vs PDF"):
                    st.dataframe(key_compare.sort_values("DABG_MATCH_KEY"), use_container_width=True)

                    st.download_button(
                        "⬇️ Download DABG Key Compare",
                        to_excel_bytes(key_compare),
                        "DABG_Key_Compare.xlsx",
                    )

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
