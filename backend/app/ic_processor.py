"""
IC Matching Report — Importable Processing Module
===================================================
Refactored from AI/ic_matching_v3.py for use as a backend service.
Enforces exactly 71 columns to match the master layout.
Computes Variance and Total values in Python (not Excel formulas)
so they are always present in the output file.
Supports reverse entity/partner matching for partner-side columns.
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

J_ENTITY = 3; J_ACCT = 4; J_ICP = 5; J_DEBIT = 15; J_CREDIT = 16

# ── Exact 71-column layout definitions from Master Template ────────────────────
FIXED_ENT_COLS = [
    ("165003", "165003:Advances provided for related parties", "S1", "Entity"),
    ("165004", "165004:Payments to subcontractors & suppliers on behalf of other parties", "S1", "Entity"),
    ("165005", "165005:Retention receivable-Intercompany (Current)", "S1", "Entity"),
    ("189001", "189001:Inter company account - receivables", "S1", "Entity"),
    ("189014", "189014:Discount on intercompany loan", "S1", "Entity"),
    ("189501", "189501:Due from Related Party - Non-current", "S1", "Entity"),
    ("224000", "224000:Due to related parties", "S2", "Partner"),
    ("224003", "224003:Advances provided from related parties", "S2", "Partner"),
    ("224009", "224009:Inter company accounts - Payables", "S2", "Partner"),
]

FIXED_PAR_COLS = [
    ("165003", "165003:Advances provided for related parties", "S1", "Partner"),
    ("165004", "165004:Payments to subcontractors & suppliers on behalf of other parties", "S1", "Partner"),
    ("165005", "165005:Retention receivable-Intercompany (Current)", "S1", "Partner"),
    ("189001", "189001:Inter company account - receivables", "S1", "Partner"),
    ("189014", "189014:Discount on intercompany loan", "S1", "Partner"),
    ("189501", "189501:Due from Related Party - Non-current", "S1", "Partner"),
    ("224000", "224000:Due to related parties", "S2", "Entity"),
    ("224003", "224003:Advances provided from related parties", "S2", "Entity"),
    ("224009", "224009:Inter company accounts - Payables", "S2", "Entity"),
]

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
    """ICP_001051 → 001051"""
    s = str(icp_code or "").strip()
    if s.startswith("ICP_"):
        return s[4:]
    return s

def extract_account_code(raw: str) -> str:
    raw = str(raw or "").strip()
    m = re.search(r"\]\.\[?(\w+)\]?:", raw)
    if m:
        val = m.group(1)
        if not val.isdigit():
            m2 = re.search(r"]:(\d{6}):", raw)
            if m2: return m2.group(1)
        return val
    m = re.match(r"(\d{6})", raw)
    return m.group(1) if m else ""

def extract_entity_code_icm(raw: str) -> str:
    m = re.match(r"(\d{6})", str(raw or "").strip())
    return m.group(1) if m else ""

def extract_entity_code_journal(raw: str) -> str:
    raw = str(raw or "").strip()
    m = re.match(r"(E?\d+\w*)", raw)
    return m.group(1) if m else ""

def to_float(val) -> float:
    if val is None or str(val).strip() in ("", " "):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def is_detail_row(vals: list) -> bool:
    entity = str(vals[J_ENTITY - 1] or "").strip()
    acct   = str(vals[J_ACCT - 1]   or "").strip()
    icp    = str(vals[J_ICP - 1]    or "").strip()
    return bool(entity or acct or icp)

def classify_account(code: str) -> str:
    code = str(code or "").strip()
    if not code or not code[0].isdigit(): return "EXTRA"
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

def read_icm_headers(ws):
    """ Builds a map of (code, tag) -> col_index from the input ICM file """
    headers = [cell.value for cell in ws[ICM_INPUT_HEADER_ROW]]
    col_map = {}
    for i, h in enumerate(headers, start=1):
        hs = str(h or "").strip()
        code = extract_account_code(hs)
        if not code: continue
        tag = "Partner" if re.search(r"\bPartner\b", hs) else "Entity"
        col_map[(code, tag)] = i
    return col_map

def read_icm_data(ws):
    data_rows = []
    for row_num, row in enumerate(
            ws.iter_rows(min_row=ICM_INPUT_DATA_START, max_row=ws.max_row), start=ICM_INPUT_DATA_START):
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
    lines = []
    
    for row_num, row in enumerate(ws.iter_rows(min_row=JOURNAL_DATA_START, max_row=ws.max_row), start=JOURNAL_DATA_START):
        vals = [cell.value for cell in row]
        label = str(vals[0] or "").strip()
        if label == "Grand Total": break
        if not is_detail_row(vals): continue
        
        entity_raw   = str(vals[J_ENTITY - 1] or "").strip()
        account_raw  = str(vals[J_ACCT - 1]   or "").strip()
        icp_raw      = str(vals[J_ICP - 1]    or "").strip()
        entity_code  = extract_entity_code_journal(entity_raw)
        account_code = extract_account_code(account_raw)
        icp_code     = extract_icp_code(icp_raw)
        debit        = to_float(vals[J_DEBIT - 1])
        credit       = to_float(vals[J_CREDIT - 1])

        if plug_mapping:
            plug_code = plug_mapping.get("plug_code")
            elim_codes = plug_mapping.get("elim_codes", set())
            if account_code in elim_codes or not account_code[:1].isdigit():
                # Map elimination accounts AND non-numeric codes (e.g. Plug_InvSh)
                # to the plug account code (188800)
                account_code = plug_code

        lines.append({
            "label": label,
            "entity_code": entity_code,
            "is_group": bool(re.match(r"^E\d+", entity_code)),
            "account_code": account_code,
            "icp_code": icp_code,
            "debit": debit,
            "credit": credit,
        })

    primary_lookup  = defaultdict(list)
    fallback_lookup = defaultdict(list)
    for line in lines:
        if not line["is_group"]:
            primary_lookup[(line["entity_code"], line["icp_code"], line["account_code"])].append(line)
        else:
            fallback_lookup[(line["icp_code"], line["account_code"])].append(line)
            
    return primary_lookup, fallback_lookup


def match_journal_to_icm(data_rows, primary_lookup, fallback_lookup):
    """
    Match journal entries to ICM rows from both entity and partner perspectives.
    Returns: (entity_updates, partner_updates)
      entity_updates  — journal Entity = ICM Entity, journal ICP = ICM Partner
      partner_updates — journal Entity = ICM Partner's 6-digit, journal ICP = ICP_ + ICM Entity
    """
    entity_updates = {}
    partner_updates = {}

    for icm_row in data_rows:
        ent_code     = icm_row["entity_code"]
        partner_code = icm_row["partner_code"]
        if not partner_code:
            continue

        # Reverse-lookup keys for partner-side
        partner_6digit = extract_6digit_from_icp(partner_code)
        entity_as_icp  = f"ICP_{ent_code}"

        # Entity-side primary: journal Entity = ICM Entity, journal ICP = ICM Partner
        for (e, p, a), jlines in primary_lookup.items():
            if e == ent_code and p == partner_code:
                net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
                if net != 0:
                    entity_updates[(ent_code, partner_code, a)] = net

        # Partner-side primary: journal Entity = Partner's 6-digit, journal ICP = ICP_Entity
        for (e, p, a), jlines in primary_lookup.items():
            if e == partner_6digit and p == entity_as_icp:
                net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
                if net != 0:
                    partner_updates[(ent_code, partner_code, a)] = net

        # Entity-side fallback (E-prefix group entities)
        for (p, a), jlines in fallback_lookup.items():
            if p == partner_code and (ent_code, partner_code, a) not in entity_updates:
                net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
                if net != 0:
                    entity_updates[(ent_code, partner_code, a)] = ("GROUP", net)

        # Partner-side fallback
        for (p, a), jlines in fallback_lookup.items():
            if p == entity_as_icp and (ent_code, partner_code, a) not in partner_updates:
                net = sum(apply_sign(j["debit"], j["credit"], a) for j in jlines)
                if net != 0:
                    partner_updates[(ent_code, partner_code, a)] = ("GROUP", net)

    return entity_updates, partner_updates

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

def write_output(ws_icm_source, data_rows, icm_header_map, updates_list, output_path, plug_code=None):
    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "ICM Matched"

    TOTAL_COLS = 71
    spacer_cols = {24, 46, 68, 70}
    for i in range(TOTAL_COLS):
        col_num = i + 1
        ws.column_dimensions[get_column_letter(col_num)].width = 4 if col_num in spacer_cols else 17.5714

    ws.row_dimensions[1].height  = 27.0
    for r in list(range(10, 20)) + [23, 28]:
        ws.row_dimensions[r].height = 14.45
    ws.row_dimensions[ICM_OUTPUT_HEADER_ROW].height = 78.75

    # Section Labels
    def _section_label(start, end, text):
        ws.merge_cells(start_row=SECTION_LABEL_ROW, start_column=start, end_row=SECTION_LABEL_ROW, end_column=end)
        c = ws.cell(SECTION_LABEL_ROW, start)
        c.value = text; c.fill = SECTION_FILL; c.font = SECTION_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")

    _section_label(25, 45, "Parent Input")
    _section_label(47, 67, "Contribution Input")
    _section_label(69, 71, "Plug Account")

    # Header Row 32
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 1), "Entity")
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 2), "Partner")

    def _write_block_headers(ent_start, par_start, var1_col, var2_col, total_col):
        for i, (code, desc, _, tag) in enumerate(FIXED_ENT_COLS):
            _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, ent_start + i), f"{code} - {desc} {tag}")
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, var1_col), "Variance")

        for i, (code, desc, _, tag) in enumerate(FIXED_PAR_COLS):
            _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, par_start + i), f"{code} - {desc} {tag}")
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, var2_col), "Variance")
        _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, total_col), "Total")

    _write_block_headers(3, 13, 12, 22, 23)
    _write_block_headers(25, 35, 34, 44, 45)
    _write_block_headers(47, 57, 56, 66, 67)

    plug_label = f"{plug_code}:Intercompany Balances Plug A/c" if plug_code else "188800:Intercompany Balances Plug A/c"
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 69), plug_label)
    _style_header_cell(ws.cell(ICM_OUTPUT_HEADER_ROW, 71), "Total")

    def _icm_num(src_r, code, tag):
        col_idx = icm_header_map.get((code, tag))
        if not col_idx: return None
        v = ws_icm_source.cell(src_r, col_idx).value
        try: return float(v) if v not in (None, "", " ") else None
        except: return None

    def _extract_val(raw):
        """Extract numeric value from update entry (handles GROUP tuples)."""
        if raw is None:
            return None, NO_FILL
        if isinstance(raw, tuple):
            return raw[1], GROUP_FILL
        return raw, MATCH_FILL

    # Unpack updates_list: each entry is (entity_updates, partner_updates)
    parent_ent, parent_par       = updates_list[0] if len(updates_list) > 0 else ({}, {})
    contrib_ent, contrib_par     = updates_list[1] if len(updates_list) > 1 else ({}, {})
    plug_ent, _                  = updates_list[2] if len(updates_list) > 2 else ({}, {})

    # S1/S2 index ranges within FIXED_ENT_COLS and FIXED_PAR_COLS
    s1_ent_codes = [c[0] for c in FIXED_ENT_COLS if c[2] == "S1"]  # first 6
    s2_par_codes = [c[0] for c in FIXED_ENT_COLS if c[2] == "S2"]  # last 3
    s1_par_codes = [c[0] for c in FIXED_PAR_COLS if c[2] == "S1"]  # first 6
    s2_ent_codes = [c[0] for c in FIXED_PAR_COLS if c[2] == "S2"]  # last 3

    for out_idx, icm_row in enumerate(data_rows):
        src_r = icm_row["row_num"]
        out_r = ICM_OUTPUT_DATA_START + out_idx
        ent   = icm_row["entity_code"]
        prt   = icm_row["partner_code"]

        _style_id_cell(ws.cell(out_r, 1), icm_row["entity"])
        _style_id_cell(ws.cell(out_r, 2), icm_row["partner"])

        def _write_block(b_ent_s, b_par_s, b_var1, b_var2, b_total,
                         is_base, ent_updates=None, par_updates=None):
            """Write one block of columns and return the computed Total value."""
            ent_vals = {}  # code → numeric value written to entity-side
            par_vals = {}  # code → numeric value written to partner-side

            # ── Entity-side columns ──────────────────────────────────────
            for i, (code, _, series, tag) in enumerate(FIXED_ENT_COLS):
                if is_base:
                    v = _icm_num(src_r, code, tag)
                    _style_data_cell(ws.cell(out_r, b_ent_s + i), v)
                    ent_vals[code] = v
                else:
                    raw = ent_updates.get((ent, prt, code)) if ent_updates else None
                    v, fill = _extract_val(raw)
                    _style_data_cell(ws.cell(out_r, b_ent_s + i), v, fill if v is not None else None)
                    ent_vals[code] = v

            # ── Variance 1 (Python-computed) ─────────────────────────────
            sum_s1_ent = sum(to_float(ent_vals.get(c)) for c in s1_ent_codes)
            sum_s2_par = sum(to_float(ent_vals.get(c)) for c in s2_par_codes)
            var1 = sum_s1_ent - sum_s2_par
            _style_data_cell(ws.cell(out_r, b_var1), var1, VARIANCE_FILL)
            ws.cell(out_r, b_var1).number_format = NUM_FORMAT

            # ── Partner-side columns ─────────────────────────────────────
            for i, (code, _, series, tag) in enumerate(FIXED_PAR_COLS):
                if is_base:
                    v = _icm_num(src_r, code, tag)
                    _style_data_cell(ws.cell(out_r, b_par_s + i), v)
                    par_vals[code] = v
                else:
                    raw = par_updates.get((ent, prt, code)) if par_updates else None
                    v, fill = _extract_val(raw)
                    _style_data_cell(ws.cell(out_r, b_par_s + i), v, fill if v is not None else None)
                    par_vals[code] = v

            # ── Variance 2 (Python-computed) ─────────────────────────────
            sum_s1_par = sum(to_float(par_vals.get(c)) for c in s1_par_codes)
            sum_s2_ent = sum(to_float(par_vals.get(c)) for c in s2_ent_codes)
            var2 = sum_s1_par - sum_s2_ent
            _style_data_cell(ws.cell(out_r, b_var2), var2, VARIANCE_FILL)
            ws.cell(out_r, b_var2).number_format = NUM_FORMAT

            # ── Total (Python-computed) ──────────────────────────────────
            total = var1 + var2
            _style_data_cell(ws.cell(out_r, b_total), total, TOTAL_FILL)
            ws.cell(out_r, b_total).number_format = NUM_FORMAT

            return total

        # Write the three main blocks and track totals
        base_total = _write_block(3, 13, 12, 22, 23, is_base=True)
        parent_total = _write_block(25, 35, 34, 44, 45, is_base=False,
                                    ent_updates=parent_ent, par_updates=parent_par)
        contrib_total = _write_block(47, 57, 56, 66, 67, is_base=False,
                                     ent_updates=contrib_ent, par_updates=contrib_par)

        # ── Plug Account block ───────────────────────────────────────────
        plug_val = 0.0
        if plug_code:
            raw = plug_ent.get((ent, prt, plug_code))
            v, fill = _extract_val(raw)
            if v is not None:
                plug_val = to_float(v)
                _style_data_cell(ws.cell(out_r, 69), v, fill)
            else:
                _style_data_cell(ws.cell(out_r, 69), None)
        else:
            _style_data_cell(ws.cell(out_r, 69), None)

        # ── Final Total = Base + Parent + Contribution + Plug ────────────
        final_total = base_total + parent_total + contrib_total + plug_val
        _style_data_cell(ws.cell(out_r, 71), final_total, TOTAL_FILL)
        ws.cell(out_r, 71).number_format = NUM_FORMAT

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    out_wb.save(output_path)
    logger.info("Saved output with computed values to %s", output_path)

def process_icm_report(icm_path, journal_paths, output_path, report_inputs_path=None):
    logger.info("=" * 65)
    logger.info("  IC Matching — 71-Column with Python-Computed Values")
    logger.info("=" * 65)

    wb_icm = openpyxl.load_workbook(icm_path, data_only=True)
    candidates = []
    for sheet in wb_icm.worksheets:
        c1 = str(sheet.cell(ICM_INPUT_HEADER_ROW, 1).value or "").strip()
        c2 = str(sheet.cell(ICM_INPUT_HEADER_ROW, 2).value or "").strip()
        if c1.lower() == "entity" and c2.lower() == "partner":
            candidates.append(sheet)

    ws_icm = min(candidates, key=lambda s: s.max_column) if candidates else wb_icm.active

    icm_header_map = read_icm_headers(ws_icm)
    data_rows = read_icm_data(ws_icm)

    plug_mapping = None; plug_code = None
    if report_inputs_path:
        plug_mapping = parse_report_inputs(report_inputs_path)
        if plug_mapping: plug_code = plug_mapping.get("plug_code")

    journal_order = ["parent_journal", "contribution_journal", "plugaccount_journal"]
    updates_list = []  # list of (entity_updates, partner_updates) tuples

    for jkey in journal_order:
        jpath = journal_paths.get(jkey)
        if not jpath:
            updates_list.append(({}, {}))
            continue

        jmap = plug_mapping if jkey == "plugaccount_journal" else None
        primary, fallback = read_journal_report(jpath, jmap)
        ent_upd, par_upd = match_journal_to_icm(data_rows, primary, fallback)
        updates_list.append((ent_upd, par_upd))

    write_output(ws_icm, data_rows, icm_header_map, updates_list, output_path, plug_code)
    return output_path
