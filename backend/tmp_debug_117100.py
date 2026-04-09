"""Debug: check what parent journal keys exist for entity 117100."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from app.ic_processor import read_journal_report, detect_icm_header_row, read_icm_data
import openpyxl

parent = r"g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx"
lookup = read_journal_report(parent)

out = open("tmp_117100_keys.txt", "w", encoding="utf-8")
out.write("=== PARENT JOURNAL KEYS for entity 117100 ===\n")
for (e, p, a), lines in sorted(lookup.items()):
    if e == "117100" or p == "117100":
        out.write(f"  KEY: ({e}, {p}, {a}) -> {len(lines)} lines\n")

out.write("\n=== ICM DATA ROWS for 117100 ===\n")
icm_path = r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx"
wb = openpyxl.load_workbook(icm_path, data_only=True)
for sheet in wb.worksheets:
    hdr, ds = detect_icm_header_row(sheet)
    data_rows = read_icm_data(sheet, ds)
    for r in data_rows:
        if r["entity_code"] == "117100" or r.get("partner_num", "") == "117100":
            out.write(f"  entity_code={r['entity_code']} partner_num={r.get('partner_num','')} entity={r['entity'][:50]} partner={r['partner'][:50]}\n")

out.write("\n=== SYNTHETIC ROW MATCHING ===\n")
# The synthetic row for E117100->ICP_007009 would have:
# ent=117100, prt_num=007009
# Entity-side key: (117100, 007009, acct) -> check if this key exists in lookup
for acct in ["534018", "433002"]:
    key = ("117100", "007009", acct)
    if key in lookup:
        out.write(f"  FOUND entity-side key {key}\n")
    else:
        out.write(f"  MISSING entity-side key {key}\n")
    rev_key = ("007009", "117100", acct)
    if rev_key in lookup:
        out.write(f"  FOUND partner-side key {rev_key}\n")
    else:
        out.write(f"  MISSING partner-side key {rev_key}\n")

out.close()
print("Done")
