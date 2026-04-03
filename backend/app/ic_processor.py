"""
IC Matching Report — Importable Processing Module
===================================================
Refactored from AI/ic_matching_v3.py for use as a backend service.

Key design:
- Output columns are derived DYNAMICALLY from the ICM source file accounts
  and the elimination accounts in report_inputs.  NO journal-only accounts
  are added as columns.
- Missing Entity–Partner rows are added only when the journal account
  exists in the ICM-defined account set.
- Parent and Contribution blocks each have their own Plug Account column.
- Computes Variance and Total values in Python (not Excel formulas).
"""

import re
import os
import logging
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

logger = logging.getLogger(__name__)

# ── Row / column constants ─────────────────────────────────────────────────────
ICM_INPUT_HEADER_ROW  = 4
ICM_INPUT_DATA_START  = 5
ICM_OUTPUT_HEADER_ROW = 32
ICM_OUTPUT_DATA_START = 33
JOURNAL_DATA_START    = 31
SECTION_LABEL_ROW     = 29



# ── Styling ───────────────────────────────────────────────────────────────────
HEADER_FONT    = Font(bold=True, color="FFFFFF", size=9)
HEADER_FILL    = PatternFill(start_color="003366", end_color="003366", fill_type="solid")
ID_FILL        = PatternFill(start_color="C8DCF0", end_color="C8DCF0", fill_type="solid")
VARIANCE_FILL  = PatternFill(start_color="DCDCDC", end_color="DCDCDC", fill_type="solid")
TOTAL_FILL     = PatternFill(start_color="C8C8C8", end_color="C8C8C8", fill_type="solid")
MATCH_FILL     = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
GROUP_FILL     = PatternFill(start_color="FFD9D9", end_color="FFD9D9", fill_type="solid")
SECTION_FILL   = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
NO_FILL        = PatternFill(fill_type=None)
ID_FONT        = Font(bold=True, size=11)
DATA_FONT      = Font(size=11)
SECTION_FONT   = Font(bold=True, color="FFFFFF", size=11)
NUM_FORMAT     = r"###,##0;\-###,##0"

DATA_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin")
)
THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)

# ═══════════════════════════════════════════════════════════════════════════
# UTILS
# ═══════════════════════════════════════════════════════════════════════════

def extract_icp_code(raw: str) -> str:
    m = re.match(r"(ICP_\w+)", str(raw or "").strip())
    return m.group(1) if m else ""

def extract_6digit_from_icp(icp_code: str) -> str:
    s = str(icp_code or "").strip()
    return s[4:] if s.startswith("ICP_") else s

def extract_account_code(raw: str) -> str:
    raw = str(raw or "").strip()
    m = re.search(r"]\.\[?(\w+)\]?:", raw)
    if m:
        val = m.group(1)
        if not val.isdigit():
            m2 = re.search(r"]:(\d{6}):", raw)
            if m2: return m2.group(1)
        return val
    m = re.match(r"(\d{6})", raw)
    return m.group(1) if m else ""

def extract_entity_code_icm(raw: str) -> str:
    """Extract and normalize entity code from ICM source data.
    Handles both pure numeric (001001) and E-prefixed (E101000) entities."""
    raw = str(raw or "").strip()
    # Try pure 6-digit first
    m = re.match(r"(\d{6})", raw)
    if m:
        return m.group(1)
    # Try E-prefixed entity
    m = re.match(r"E(\d+)", raw)
    if m:
        return m.group(1)  # Strip the E prefix for normalized matching
    return ""

def extract_entity_code_journal(raw: str) -> str:
    raw = str(raw or "").strip()
    m = re.match(r"(E?\d+\w*)", raw)
    return m.group(1) if m else ""

def normalize_entity_code(code: str) -> str:
    """Normalize entity codes to numeric-only format for matching.
    Strips 'E' prefix so journal codes ('E101000') match ICM codes ('101000')."""
    code = str(code or "").strip()
    if code.startswith("E") and len(code) > 1 and code[1:].replace("_", "").isdigit():
        return code[1:]
    return code

def to_float(val) -> float:
    if val is None or str(val).strip() in ("", " "):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def get_journal_indices(ws):
    header_row = JOURNAL_DATA_START - 1
    headers = [str(c.value or "").strip().lower() for c in ws[header_row]]
    col_map = {}
    for i, h in enumerate(headers):
        if h: col_map[h] = i
    return {
        "entity": col_map.get('entity', 2),
        "acct": col_map.get('account', 3),
        "icp": col_map.get('intercompany', 4),
        "debit": col_map.get('debit', 14),
        "credit": col_map.get('credit', 15)
    }

def is_detail_row(vals: list, indices: dict) -> bool:
    try:
        entity = str(vals[indices["entity"]] or "").strip()
        acct   = str(vals[indices["acct"]]   or "").strip()
        icp    = str(vals[indices["icp"]]    or "").strip()
        return bool(entity or acct or icp)
    except IndexError:
        return False

def classify_account(code: str) -> str:
    """Classify numeric account as S1 (asset) or S2 (liability)."""
    code = str(code or "").strip()
    if not code or not code[0].isdigit():
        return "EXTRA"
    s = code[0]
    if s in ("1", "5"): return "S1"
    elif s in ("2", "3", "4"): return "S2"
    return "EXTRA"

def apply_sign(debit: float, credit: float, account_code: str) -> float:
    code = str(account_code or "").strip()
    s = code[0] if code and code[0].isdigit() else "0"
    if s in ("1", "5"):
        return debit - credit
    elif s in ("2", "3", "4"):
        return credit - debit
    return debit - credit

def parse_report_inputs(filepath: str) -> dict:
    if not filepath or not os.path.exists(filepath): return None
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = wb.active
        headers = [str(c.value or "").strip() for c in ws[1]]
        plug_code = None
        for h in headers:
            if "Plug Account:" in h:
                m = re.search(r"(\d{6})", h)
                if m: plug_code = m.group(1)
                break
        if not plug_code: return None
        elim_codes = set()
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            val = str(row[0] or "").strip()
            if not val: continue
            code = extract_account_code(val)
            if code: elim_codes.add(code)
        return {"plug_code": plug_code, "elim_codes": elim_codes}
    except Exception as e:
        logger.error("Error parsing report inputs: %s", e)
        return None

# ═══════════════════════════════════════════════════════════════════════════
# ICM PARSING
# ═══════════════════════════════════════════════════════════════════════════

def detect_icm_header_row(ws):
    """Auto-detect the row containing 'Entity' col 1 / 'Partner' col 2.
    Scans first 50 rows.  Returns (header_row, data_start_row)."""
    for r in range(1, min(51, ws.max_row + 1)):
        v1 = str(ws.cell(r, 1).value or "").strip().lower()
        v2 = str(ws.cell(r, 2).value or "").strip().lower()
        if v1 == "entity" and v2 == "partner":
            return r, r + 1
    return ICM_INPUT_HEADER_ROW, ICM_INPUT_DATA_START

def read_icm_headers(ws, header_row=None):
    """Builds a map of (code, tag) -> col_index from the ICM file header row."""
    if header_row is None:
        header_row = ICM_INPUT_HEADER_ROW
    headers = [cell.value for cell in ws[header_row]]
    col_map = {}
    for i, h in enumerate(headers, start=1):
        hs = str(h or "").strip()
        code = extract_account_code(hs)
        if not code: continue
        tag = "Partner" if re.search(r"\bPartner\b", hs) else "Entity"
        col_map[(code, tag)] = i
    return col_map

def read_icm_account_columns(ws, header_row=None):
    """Extract the ordered list of (code, description, tag) from the FIRST
    entity-side + partner-side block in the ICM source headers.

    The ICM source repeats the same block of accounts 3 times (for Base /
    Parent / Contribution).  We only need the first block to define the
    output structure.

    Returns (ent_cols, par_cols) where each is a list of
    (code, description, series, tag) tuples — one per unique account,
    preserving order."""
    if header_row is None:
        header_row = ICM_INPUT_HEADER_ROW

    # Collect all (code, desc, tag, col) from the header row
    all_headers = []
    for c in range(1, ws.max_column + 1):
        v = str(ws.cell(header_row, c).value or "").strip()
        if not v:
            continue
        code = extract_account_code(v)
        if not code:
            continue
        tag = "Partner" if re.search(r"\bPartner\b", v) else "Entity"
        all_headers.append((code, v, tag, c))

    # De-duplicate: keep the FIRST occurrence of each (code, tag) pair
    seen = set()
    ent_cols = []
    par_cols = []
    for code, desc, tag, _col in all_headers:
        key = (code, tag)
        if key in seen:
            continue
        seen.add(key)
        series = classify_account(code)
        if tag == "Entity":
            ent_cols.append((code, desc, series, tag))
        else:
            par_cols.append((code, desc, series, tag))

    return ent_cols, par_cols

def read_icm_data(ws, data_start=None):
    if data_start is None:
        data_start = ICM_INPUT_DATA_START
    data_rows = []
    for row_num, row in enumerate(
            ws.iter_rows(min_row=data_start, max_row=ws.max_row), start=data_start):
        vals = [cell.value for cell in row]
        entity_raw  = str(vals[0] or "").strip()
        partner_raw = str(vals[1] or "").strip() if len(vals) > 1 else ""
        if not entity_raw and not partner_raw: continue
        data_rows.append({
            "row_num":      row_num,
            "entity":       entity_raw,
            "partner":      partner_raw,
            "entity_code":  extract_entity_code_icm(entity_raw),
            "partner_code": extract_icp_code(partner_raw),
        })
    return data_rows

# ═══════════════════════════════════════════════════════════════════════════
# JOURNAL READING & MATCHING
# ═══════════════════════════════════════════════════════════════════════════

def read_journal_report(filepath: str, plug_mapping: dict = None):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    indices = get_journal_indices(ws)
    lines = []

    for row_num, row in enumerate(ws.iter_rows(min_row=JOURNAL_DATA_START, max_row=ws.max_row), start=JOURNAL_DATA_START):
        vals = [cell.value for cell in row]
        if not vals: continue
        label = str(vals[0] or "").strip()
        if label == "Grand Total": break
        if not is_detail_row(vals, indices): continue

        try:
            entity_raw   = str(vals[indices["entity"]] or "").strip()
            account_raw  = str(vals[indices["acct"]]   or "").strip()
            icp_raw      = str(vals[indices["icp"]]    or "").strip()
            debit        = to_float(vals[indices["debit"]])
            credit       = to_float(vals[indices["credit"]])
        except IndexError:
            continue

        entity_code_raw = extract_entity_code_journal(entity_raw)
        entity_code     = normalize_entity_code(entity_code_raw)
        account_code    = extract_account_code(account_raw)
        icp_code        = extract_icp_code(icp_raw)

        if plug_mapping:
            plug_code = plug_mapping.get("plug_code")
            elim_codes = plug_mapping.get("elim_codes", set())
            if account_code in elim_codes or not account_code[:1].isdigit():
                account_code = plug_code

        lines.append({
            "label": label,
            "entity_code": entity_code,
            "is_group": bool(re.match(r"^E\d+", entity_code_raw)),
            "account_code": account_code,
            "icp_code": icp_code,
            "debit": debit,
            "credit": credit,
        })

    primary_lookup  = defaultdict(list)
    fallback_lookup = defaultdict(list)
    for line in lines:
        if not line["entity_code"] or not line["icp_code"] or not line["account_code"]:
            continue
        if not line["is_group"]:
            primary_lookup[(line["entity_code"], line["icp_code"], line["account_code"])].append(line)
        else:
            fallback_lookup[(line["entity_code"], line["icp_code"], line["account_code"])].append(line)

    return primary_lookup, fallback_lookup


def read_journal_labels(filepath: str) -> dict:
    """code → full label string (e.g. 'E117100' → 'E117100:QD UK Holdings LP')."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    indices = get_journal_indices(ws)
    label_map = {}
    for row in ws.iter_rows(min_row=JOURNAL_DATA_START, max_row=ws.max_row):
        vals = [cell.value for cell in row]
        if not vals: continue
        if str(vals[0] or "").strip() == "Grand Total":
            break
        if not is_detail_row(vals, indices):
            continue
        try:
            entity_raw = str(vals[indices["entity"]] or "").strip()
            icp_raw    = str(vals[indices["icp"]] or "").strip()
        except IndexError:
            continue
        ent_code = extract_entity_code_journal(entity_raw)
        ent_code_norm = normalize_entity_code(ent_code)
        icp_code = extract_icp_code(icp_raw)
        if ent_code and entity_raw:
            label_map[ent_code] = entity_raw
            if ent_code_norm != ent_code:
                label_map[ent_code_norm] = entity_raw
        if icp_code and icp_raw:
            label_map[icp_code] = icp_raw
    return label_map


def match_journal_to_icm(data_rows, primary_lookup, fallback_lookup):
    """Returns (primary_updates, fallback_updates) dicts of (ent, icp, acct) → net."""
    primary_updates = {}
    fallback_updates = {}

    for (e, p, a), jlines in primary_lookup.items():
        net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
        if net != 0:
            primary_updates[(e, p, a)] = net

    for (ge, p, a), jlines in fallback_lookup.items():
        net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
        if net != 0:
            fallback_updates[(ge, p, a)] = net

    return primary_updates, fallback_updates

# ═══════════════════════════════════════════════════════════════════════════
# OUTPUT WRITER
# ═══════════════════════════════════════════════════════════════════════════

def _style_header_cell(cell, text):
    cell.value = text; cell.font = HEADER_FONT; cell.fill = HEADER_FILL; cell.border = THIN_BORDER
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def _style_id_cell(cell, text):
    cell.value = text; cell.font = ID_FONT; cell.fill = ID_FILL; cell.border = DATA_BORDER
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

def _style_data_cell(cell, value, fill=None):
    cell.value = value; cell.font = DATA_FONT; cell.fill = fill if fill is not None else NO_FILL
    cell.border = DATA_BORDER; cell.alignment = Alignment(vertical="top", wrap_text=True)
    if isinstance(value, (int, float)): cell.number_format = NUM_FORMAT


def _block_positions(start, num_accts_ent, num_accts_par, has_plug=False):
    """Compute column positions for one block.
    Returns dict with keys: ent_start, var1, par_start, var2, total, plug (or None), spacer."""
    ent_start = start
    var1      = start + num_accts_ent
    par_start = var1 + 1
    var2      = par_start + num_accts_par
    total     = var2 + 1
    if has_plug:
        plug   = total + 1
        spacer = plug + 1
    else:
        plug   = None
        spacer = total + 1
    return {
        "ent_start": ent_start, "var1": var1,
        "par_start": par_start, "var2": var2,
        "total": total, "plug": plug, "spacer": spacer,
    }


def write_output(ws_icm_source, data_rows, icm_header_map,
                 updates_list, output_path, plug_code=None,
                 ent_cols=None, par_cols=None):
    """Write the ICM output workbook.

    Parameters
    ----------
    ent_cols, par_cols : list of (code, desc, series, tag) tuples
        Defines the account columns for every block — driven by the ICM source.
    """
    if not ent_cols or not par_cols:
        raise ValueError("ent_cols and par_cols must be provided (derived from ICM source)")

    n_ent = len(ent_cols)
    n_par = len(par_cols)

    # ── Compute block positions ──────────────────────────────────────────
    blk_base = _block_positions(3, n_ent, n_par, has_plug=False)
    blk_par  = _block_positions(blk_base["spacer"] + 1, n_ent, n_par, has_plug=True)
    blk_cont = _block_positions(blk_par["spacer"] + 1, n_ent, n_par, has_plug=True)

    # Plug Account section: plug_value | Total
    plug_section_start = blk_cont["spacer"] + 1
    plug_section = {
        "plug_col": plug_section_start,
        "total": plug_section_start + 1,
        "spacer": plug_section_start + 2,
    }
    col_final = plug_section["spacer"] + 1
    total_cols = col_final

    spacer_cols = {blk_base["spacer"], blk_par["spacer"], blk_cont["spacer"], plug_section["spacer"]}

    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "ICM Matched"

    for i in range(total_cols):
        col_num = i + 1
        ws.column_dimensions[get_column_letter(col_num)].width = 4 if col_num in spacer_cols else 17.5714

    ws.row_dimensions[1].height = 27.0
    for r in list(range(10, 20)) + [23, 28]:
        ws.row_dimensions[r].height = 14.45
    ws.row_dimensions[ICM_OUTPUT_HEADER_ROW].height = 78.75

    # ── Section Labels ───────────────────────────────────────────────────
    def _section_label(start, end, text):
        ws.merge_cells(start_row=SECTION_LABEL_ROW, start_column=start,
                       end_row=SECTION_LABEL_ROW, end_column=end)
        c = ws.cell(SECTION_LABEL_ROW, start)
        c.value = text; c.fill = SECTION_FILL; c.font = SECTION_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")

    _section_label(blk_par["ent_start"], blk_par["plug"] or blk_par["total"],  "Parent Input")
    _section_label(blk_cont["ent_start"], blk_cont["plug"] or blk_cont["total"], "Contribution Input")
    _section_label(plug_section["plug_col"], plug_section["total"], "Plug Account")

    # ── Header Row ───────────────────────────────────────────────────────
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 1), "Entity")
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 2), "Partner")

    plug_label_base = f"{plug_code}:Intercompany Balances Plug A/c" if plug_code else "188800:Intercompany Balances Plug A/c"
    plug_label_parent = f"{plug_label_base}\n(Entity \u2192 Parent)"
    plug_label_contrib = f"{plug_label_base}\n(Parent \u2192 Entity)"

    def _write_block_headers(blk, has_plug_hdr=False, plug_hdr_text=None):
        for i, (code, desc, _, tag) in enumerate(ent_cols):
            _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["ent_start"] + i),
                               f"{code} - {desc} {tag}" if tag not in desc else desc)
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["var1"]), "Variance")

        for i, (code, desc, _, tag) in enumerate(par_cols):
            _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["par_start"] + i),
                               f"{code} - {desc} {tag}" if tag not in desc else desc)
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["var2"]), "Variance")
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["total"]), "Total")

        if has_plug_hdr and blk["plug"] is not None:
            _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, blk["plug"]),
                               plug_hdr_text or plug_label_base)

    _write_block_headers(blk_base, has_plug_hdr=False)
    _write_block_headers(blk_par, has_plug_hdr=True, plug_hdr_text=plug_label_parent)
    _write_block_headers(blk_cont, has_plug_hdr=True, plug_hdr_text=plug_label_contrib)

    # Plug Account section headers
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, plug_section["plug_col"]), plug_label_base)
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, plug_section["total"]), "Total")

    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, col_final), "Final Total")

    # ── Scale factor for output values ────────────────────────────────────
    SCALE = 1.0

    # ── Helper: read base value from ICM source ──────────────────────────
    def _icm_num(src_r, code, tag):
        if src_r is None: return None
        col_idx = icm_header_map.get((code, tag))
        if not col_idx: return None
        v = ws_icm_source.cell(src_r, col_idx).value
        try:
            val = float(v) if v not in (None, "", " ") else None
            return val * SCALE if val is not None else None
        except: return None

    def _extract_val(raw):
        if raw is None:
            return None, NO_FILL
        if isinstance(raw, tuple):
            return raw[1] * SCALE, GROUP_FILL
        return raw * SCALE, MATCH_FILL

    # ── Unpack updates ───────────────────────────────────────────────────
    def _unpack(idx):
        entry = updates_list[idx] if len(updates_list) > idx else ({}, {})
        return entry if isinstance(entry, tuple) else (entry, {})

    parent_primary,  parent_fallback  = _unpack(0)
    contrib_primary, contrib_fallback = _unpack(1)
    plug_primary,    plug_fallback    = _unpack(2)

    def _merge_fallback(primary, fallback):
        merged = dict(primary)
        for key, net in fallback.items():
            if key not in merged:
                merged[key] = (key[0], net)
        return merged

    parent_primary  = _merge_fallback(parent_primary,  parent_fallback)
    contrib_primary = _merge_fallback(contrib_primary, contrib_fallback)
    plug_primary    = _merge_fallback(plug_primary,    plug_fallback)

    # S1/S2 groups for variance
    s1_ent_codes = [c[0] for c in ent_cols if c[2] == "S1"]
    s2_ent_codes = [c[0] for c in ent_cols if c[2] == "S2"]
    s1_par_codes = [c[0] for c in par_cols if c[2] == "S1"]
    s2_par_codes = [c[0] for c in par_cols if c[2] == "S2"]

    # Consumed sets — ONE set per journal to ensure each value appears only ONCE
    consumed_parent  = set()
    consumed_contrib = set()
    consumed_plug    = set()

    # ── Write data rows ──────────────────────────────────────────────────
    for out_idx, icm_row in enumerate(data_rows):
        src_r = icm_row["row_num"]
        out_r = ICM_OUTPUT_DATA_START + out_idx
        ent   = icm_row["entity_code"]
        prt   = icm_row["partner_code"]

        _style_id_cell(ws.cell(out_r, 1), icm_row["entity"])
        _style_id_cell(ws.cell(out_r, 2), icm_row["partner"])

        # Direct match only: Journal Entity → ICM Entity, Journal ICP → ICM Partner
        # No fabricated ICP codes — entity names never start with ICP_
        can_match = bool(ent and prt)

        # Reverse key components for partner-to-entity matching
        prt_entity = normalize_entity_code(extract_6digit_from_icp(prt)) if prt else ""
        reverse_icp_plain = f"ICP_{ent}" if ent else ""
        reverse_icp_e = f"ICP_E{ent}" if ent else ""
        can_reverse = bool(prt_entity and ent)

        def _write_block(blk, is_base, primary=None, consumed=None):
            """Write entity-side, partner-side, variances, total for one block.
            Journal matching uses DIRECT key (ent, prt, code) for ALL columns."""
            ent_vals = {}
            par_vals = {}

            # ── Entity-side (Direct Match) ──────────────────────────────
            for i, (code, _, series, tag) in enumerate(ent_cols):
                if is_base:
                    v = _icm_num(src_r, code, tag)
                    _style_data_cell(ws.cell(out_r, blk["ent_start"] + i), v)
                    ent_vals[code] = v
                else:
                    pri_key = (ent, prt, code)
                    raw = None
                    if can_match and primary and pri_key in primary and \
                       (consumed is None or pri_key not in consumed):
                        raw = primary[pri_key]
                        if consumed is not None:
                            consumed.add(pri_key)
                    v, fill = _extract_val(raw)
                    _style_data_cell(ws.cell(out_r, blk["ent_start"] + i), v,
                                     fill if v is not None else None)
                    ent_vals[code] = v

            # ── Variance 1 ──────────────────────────────────────────────
            sum_s1 = sum(to_float(ent_vals.get(c)) for c in s1_ent_codes)
            sum_s2 = sum(to_float(ent_vals.get(c)) for c in s2_ent_codes)
            var1 = sum_s1 - sum_s2
            _style_data_cell(ws.cell(out_r, blk["var1"]), var1, VARIANCE_FILL)
            ws.cell(out_r, blk["var1"]).number_format = NUM_FORMAT

            # ── Partner-side (Reverse key: partner as entity, entity as ICP) ──
            for i, (code, _, series, tag) in enumerate(par_cols):
                if is_base:
                    v = _icm_num(src_r, code, tag)
                    _style_data_cell(ws.cell(out_r, blk["par_start"] + i), v)
                    par_vals[code] = v
                else:
                    raw_par = None
                    if can_match and can_reverse and primary:
                        for rk in ((prt_entity, reverse_icp_plain, code),
                                   (prt_entity, reverse_icp_e, code)):
                            if rk in primary and (consumed is None or rk not in consumed):
                                raw_par = primary[rk]
                                if consumed is not None:
                                    consumed.add(rk)
                                break
                    v, fill = _extract_val(raw_par)
                    _style_data_cell(ws.cell(out_r, blk["par_start"] + i), v,
                                     fill if v is not None else None)
                    par_vals[code] = v

            # ── Variance 2 ──────────────────────────────────────────────
            sum_s1p = sum(to_float(par_vals.get(c)) for c in s1_par_codes)
            sum_s2p = sum(to_float(par_vals.get(c)) for c in s2_par_codes)
            var2 = sum_s1p - sum_s2p
            _style_data_cell(ws.cell(out_r, blk["var2"]), var2, VARIANCE_FILL)
            ws.cell(out_r, blk["var2"]).number_format = NUM_FORMAT

            # ── Total ───────────────────────────────────────────────────
            total = var1 + var2
            _style_data_cell(ws.cell(out_r, blk["total"]), total, TOTAL_FILL)
            ws.cell(out_r, blk["total"]).number_format = NUM_FORMAT

            return total

        # ── Write blocks ─────────────────────────────────────────────────
        base_total = _write_block(blk_base, is_base=True)

        parent_total = _write_block(blk_par, is_base=False,
                                    primary=parent_primary,
                                    consumed=consumed_parent)

        contrib_total = _write_block(blk_cont, is_base=False,
                                     primary=contrib_primary,
                                     consumed=consumed_contrib)

        # ── Plug Account — Parent block ──────────────────────────────────
        plug_par_val = 0.0
        if plug_code and ent and prt and blk_par["plug"]:
            plug_key = (ent, prt, plug_code)
            raw = None
            if plug_key in parent_primary and plug_key not in consumed_parent:
                raw = parent_primary[plug_key]
                consumed_parent.add(plug_key)
            v, fill = _extract_val(raw)
            if v is not None:
                plug_par_val = to_float(v)
                _style_data_cell(ws.cell(out_r, blk_par["plug"]), v, fill)
            else:
                _style_data_cell(ws.cell(out_r, blk_par["plug"]), None)

        # ── Plug Account — Contribution block (reverse: partner→entity) ──
        plug_cont_val = 0.0
        if plug_code and ent and prt and blk_cont["plug"]:
            raw = None
            if can_reverse:
                for rk in ((prt_entity, reverse_icp_plain, plug_code),
                           (prt_entity, reverse_icp_e, plug_code)):
                    if rk in contrib_primary and rk not in consumed_contrib:
                        raw = contrib_primary[rk]
                        consumed_contrib.add(rk)
                        break
            v, fill = _extract_val(raw)
            if v is not None:
                plug_cont_val = to_float(v)
                _style_data_cell(ws.cell(out_r, blk_cont["plug"]), v, fill)
            else:
                _style_data_cell(ws.cell(out_r, blk_cont["plug"]), None)

        # ── Plug Account Section (Standalone) ────────────────────────────
        plug_val = 0.0
        if plug_code and ent and prt:
            # Direct key only — no vice-versa
            plug_key = (ent, prt, plug_code)
            raw = None
            if plug_key in plug_primary and plug_key not in consumed_plug:
                raw = plug_primary[plug_key]
                consumed_plug.add(plug_key)
            v, fill = _extract_val(raw)
            if v is not None:
                plug_val = to_float(v)
                _style_data_cell(ws.cell(out_r, plug_section["plug_col"]), v, fill)
            else:
                _style_data_cell(ws.cell(out_r, plug_section["plug_col"]), None)
        else:
            _style_data_cell(ws.cell(out_r, plug_section["plug_col"]), None)

        # Plug Section Total
        _style_data_cell(ws.cell(out_r, plug_section["total"]),
                         plug_val if plug_val else None, TOTAL_FILL)
        ws.cell(out_r, plug_section["total"]).number_format = NUM_FORMAT

        # ── Final Total ──────────────────────────────────────────────────
        final_total = base_total + parent_total + contrib_total + plug_par_val + plug_cont_val + plug_val
        _style_data_cell(ws.cell(out_r, col_final), final_total, TOTAL_FILL)
        ws.cell(out_r, col_final).number_format = NUM_FORMAT

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    out_wb.save(output_path)
    logger.info("Saved output to %s  (%d rows, %d cols)", output_path, len(data_rows), total_cols)


# ═══════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

def process_icm_report(icm_path, journal_paths, output_path, report_inputs_path=None):
    logger.info("=" * 65)
    logger.info("  IC Matching — Dynamic ICM-Driven Columns")
    logger.info("=" * 65)

    wb_icm = openpyxl.load_workbook(icm_path, data_only=True)

    # ── Auto-detect sheet and header row ─────────────────────────────────
    candidates = []
    for sheet in wb_icm.worksheets:
        hdr_row, data_start = detect_icm_header_row(sheet)
        c1 = str(sheet.cell(hdr_row, 1).value or "").strip().lower()
        c2 = str(sheet.cell(hdr_row, 2).value or "").strip().lower()
        if c1 == "entity" and c2 == "partner":
            candidates.append((sheet, hdr_row, data_start))

    if candidates:
        ws_icm, icm_hdr_row, icm_data_start = min(candidates, key=lambda t: t[0].max_column)
    else:
        ws_icm = wb_icm.active
        icm_hdr_row, icm_data_start = detect_icm_header_row(ws_icm)

    logger.info("Using sheet '%s' header=%d data=%d", ws_icm.title, icm_hdr_row, icm_data_start)

    # ── Read ICM account columns (dynamic) ───────────────────────────────
    icm_header_map = read_icm_headers(ws_icm, header_row=icm_hdr_row)
    ent_cols_from_icm, par_cols_from_icm = read_icm_account_columns(ws_icm, header_row=icm_hdr_row)

    # Also read from the wider sheet if available — to capture all accounts
    if len(candidates) > 1:
        ws_wide = max(candidates, key=lambda t: t[0].max_column)[0]
        wide_hdr = max(candidates, key=lambda t: t[0].max_column)[1]
        ent_wide, par_wide = read_icm_account_columns(ws_wide, header_row=wide_hdr)
        icm_header_map_wide = read_icm_headers(ws_wide, header_row=wide_hdr)
        # Merge: add any accounts from wider sheet not already present
        existing_ent = {c[0] for c in ent_cols_from_icm}
        existing_par = {c[0] for c in par_cols_from_icm}
        for col in ent_wide:
            if col[0] not in existing_ent:
                ent_cols_from_icm.append(col)
                existing_ent.add(col[0])
        for col in par_wide:
            if col[0] not in existing_par:
                par_cols_from_icm.append(col)
                existing_par.add(col[0])
        icm_header_map.update(icm_header_map_wide)

    # ── Read elimination accounts from report_inputs ─────────────────────
    plug_mapping = None
    plug_code = None
    if report_inputs_path:
        plug_mapping = parse_report_inputs(report_inputs_path)
        if plug_mapping:
            plug_code = plug_mapping.get("plug_code")

    # Merge elimination accounts into the column lists if not already present
    if plug_mapping:
        elim_codes = plug_mapping.get("elim_codes", set())
        existing_ent_codes = {c[0] for c in ent_cols_from_icm}
        existing_par_codes = {c[0] for c in par_cols_from_icm}
        for ec in sorted(elim_codes):
            if not ec or not ec[0].isdigit():
                continue
            series = classify_account(ec)
            if series == "EXTRA":
                series = "S1"  # Default for accounts starting with 6-9
            if ec not in existing_ent_codes:
                # Entity-side: S1 → Entity tag, S2 → Partner tag
                ent_tag = "Entity" if series == "S1" else "Partner"
                ent_cols_from_icm.append((ec, f"{ec}:{ec}", series, ent_tag))
                existing_ent_codes.add(ec)
            if ec not in existing_par_codes:
                # Partner-side: reversed tags
                par_tag = "Partner" if series == "S1" else "Entity"
                par_cols_from_icm.append((ec, f"{ec}:{ec}", series, par_tag))
                existing_par_codes.add(ec)

    # Build the set of valid account codes (ICM + elimination) for filtering
    valid_accounts = {c[0] for c in ent_cols_from_icm} | {c[0] for c in par_cols_from_icm}
    if plug_code:
        valid_accounts.add(plug_code)

    logger.info("Output columns: %d entity-side, %d partner-side",
                len(ent_cols_from_icm), len(par_cols_from_icm))
    logger.info("Valid account codes: %s", sorted(valid_accounts))

    # ── Read ICM data rows ───────────────────────────────────────────────
    data_rows = read_icm_data(ws_icm, data_start=icm_data_start)

    # ── Read journals ────────────────────────────────────────────────────
    journal_order = ["parent_journal", "contribution_journal", "plugaccount_journal"]
    updates_list = []

    for jkey in journal_order:
        jpath = journal_paths.get(jkey)
        if not jpath:
            updates_list.append(({}, {}))
            continue

        jmap = plug_mapping if jkey == "plugaccount_journal" else None
        primary, fallback = read_journal_report(jpath, jmap)

        # FILTER: only keep matches where account exists in ICM/elimination set
        primary_filtered = {k: v for k, v in primary.items() if k[2] in valid_accounts}
        fallback_filtered = {k: v for k, v in fallback.items() if k[2] in valid_accounts}

        primary_updates, fallback_updates = match_journal_to_icm(
            data_rows, primary_filtered, fallback_filtered)
        updates_list.append((primary_updates, fallback_updates))

    # ── Discover missing Entity/Partner pairs ────────────────────────────
    all_labels = {}
    for jkey in journal_order:
        jpath = journal_paths.get(jkey)
        if jpath:
            all_labels.update(read_journal_labels(jpath))

    existing_pairs = {(r["entity_code"], r["partner_code"]) for r in data_rows}

    missing_pairs = set()
    for primary_upd, fallback_upd in updates_list:
        for (ent, icp, acct) in primary_upd:
            if acct in valid_accounts and (ent, icp) not in existing_pairs:
                missing_pairs.add((ent, icp))
        for (ent, icp, acct) in fallback_upd:
            if acct in valid_accounts and (ent, icp) not in existing_pairs:
                missing_pairs.add((ent, icp))

    # ── De-duplicate reverse pairs ────────────────────────────────────
    # Journal entries come in pairs: a primary entry (entity→E-prefix partner)
    # and a fallback entry (E-prefix entity→partner). For example:
    #   Primary:  (001001, ICP_E101000, 534018)  → entity→partner direction
    #   Fallback: (101000,  ICP_001001, 534018)  → partner→entity direction
    # The forward row already shows BOTH via entity-side and partner-side columns,
    # so we remove the fallback synthetic row to avoid double-counting.
    # IMPORTANT: Only remove when one direction has ICP_E... and the other has ICP_...
    # to distinguish primary↔fallback pairs from two independent primary entries
    # (like 001032↔001033 which are both non-E-prefix).
    pairs_to_remove = set()
    for (ent, icp) in list(missing_pairs):
        # Only consider pairs where the ICP does NOT have E prefix
        # (these are fallback/reverse entries)
        if icp.startswith("ICP_E"):
            continue  # This is a primary pair — never remove
        
        # Compute the reverse pair (which would be the primary direction)
        icp_digit = icp[4:] if icp.startswith("ICP_") else icp
        rev_ent = normalize_entity_code(icp_digit)
        rev_icp_e = f"ICP_E{ent}"
        
        # Only remove if the E-prefix reverse pair exists (confirming this is
        # truly a fallback mirror, not an independent non-E pair)
        if (rev_ent, rev_icp_e) in missing_pairs:
            pairs_to_remove.add((ent, icp))
    
    missing_pairs -= pairs_to_remove
    if pairs_to_remove:
        logger.info("Removed %d reverse duplicate pairs", len(pairs_to_remove))

    for ent_code, icp_code in sorted(missing_pairs):
        data_rows.append({
            "row_num":      None,
            "entity":       all_labels.get(ent_code, ent_code),
            "partner":      all_labels.get(icp_code, icp_code),
            "entity_code":  ent_code,
            "partner_code": icp_code,
        })

    if missing_pairs:
        logger.info("Added %d synthetic rows for missing entity/partner pairs", len(missing_pairs))

    # ── Re-match after adding synthetic rows ─────────────────────────────
    updates_list_final = []
    for jkey in journal_order:
        jpath = journal_paths.get(jkey)
        if not jpath:
            updates_list_final.append(({}, {}))
            continue

        jmap = plug_mapping if jkey == "plugaccount_journal" else None
        primary, fallback = read_journal_report(jpath, jmap)

        # FILTER: only ICM-defined accounts
        primary_filtered = {k: v for k, v in primary.items() if k[2] in valid_accounts}
        fallback_filtered = {k: v for k, v in fallback.items() if k[2] in valid_accounts}

        primary_updates, fallback_updates = match_journal_to_icm(
            data_rows, primary_filtered, fallback_filtered)
        updates_list_final.append((primary_updates, fallback_updates))

    # ── Write output ─────────────────────────────────────────────────────
    write_output(ws_icm, data_rows, icm_header_map, updates_list_final, output_path,
                 plug_code, ent_cols=ent_cols_from_icm, par_cols=par_cols_from_icm)
    return output_path