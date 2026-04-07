"""Trace exactly what happens for synthetic row E117100->ICP_007009."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import openpyxl, logging
logging.basicConfig(level=logging.WARNING)

from app.ic_processor import (
    read_journal_report, match_journal_to_icm, parse_report_inputs,
    detect_icm_header_row, read_icm_data, read_icm_account_columns,
    read_icm_headers, classify_account, apply_sign
)

# Read inputs same as process_icm_report does
icm_path = r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx"
parent = r"g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx"
report_inputs = r"g:\FCCS\backend\uploads\reports\31\inputs\report Inputs.xlsx"

plug_mapping = parse_report_inputs(report_inputs)
valid_accounts = {'111006', '120030', '121015', '188600', '433002', '534018'}

out = open("tmp_trace_117100.txt", "w", encoding="utf-8")

# Read parent journal
lookup = read_journal_report(parent)
lookup_filtered = {k: v for k, v in lookup.items() if k[2] in valid_accounts}

out.write("=== FILTERED LOOKUP keys related to 117100 ===\n")
for (e, p, a), lines in sorted(lookup_filtered.items()):
    if e == "117100" or p == "117100":
        net = sum(apply_sign(l["debit"], l["credit"], a) for l in lines)
        out.write(f"  ({e}, {p}, {a}) -> net={net}\n")

# Build updates
wb = openpyxl.load_workbook(icm_path, data_only=True)
ws = wb.worksheets[0]
hdr, ds = detect_icm_header_row(ws)
data_rows = read_icm_data(ws, ds)

updates = match_journal_to_icm(data_rows, lookup_filtered)

out.write("\n=== UPDATES keys related to 117100 ===\n")
for (e, p, a), v in sorted(updates.items()):
    if e == "117100" or p == "117100":
        out.write(f"  ({e}, {p}, {a}) -> {v}\n")

# Check if key (117100, 007009, 534018) is in updates
key = ("117100", "007009", "534018")
out.write(f"\n=== KEY {key} in updates? {key in updates} ===\n")
if key in updates:
    out.write(f"  Value: {updates[key]}\n")

# Now simulate what _write_block does for a synthetic row
out.write("\n=== SIMULATING _write_block for E117100->ICP_007009 ===\n")
ent = "117100"
prt_num = "007009"
can_match = bool(ent and prt_num)
out.write(f"  ent={ent}, prt_num={prt_num}, can_match={can_match}\n")

for code in ["534018", "433002", "111006", "120030", "121015"]:
    ent_key = (ent, prt_num, code)
    rev_key = (prt_num, ent, code)
    out.write(f"\n  Account {code}:\n")
    out.write(f"    Entity-side key {ent_key}: in updates = {ent_key in updates}")
    if ent_key in updates:
        out.write(f" -> value = {updates[ent_key]}")
    out.write(f"\n    Partner-side key {rev_key}: in updates = {rev_key in updates}")
    if rev_key in updates:
        out.write(f" -> value = {updates[rev_key]}")
    out.write("\n")

out.close()
print("Done")
