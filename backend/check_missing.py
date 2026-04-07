import openpyxl, sys
sys.path.insert(0, r'G:\FCCS\backend')
from app.ic_processor import read_icm_data, detect_icm_header_row, read_journal_report
import os
base = r'G:\FCCS\backend\uploads\reports\31\inputs'
icm_path = os.path.join(base, 'IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx')
wb_icm = openpyxl.load_workbook(icm_path, data_only=True)
ws_icm = None
for s in wb_icm.worksheets:
    if s.max_row > 10: ws_icm = s; break
hdr_row, data_start = detect_icm_header_row(ws_icm)
data_rows = read_icm_data(ws_icm, data_start=data_start)
existing_pairs = {(r["entity_code"], r.get("partner_num", "")) for r in data_rows}

parent_lookup = read_journal_report(os.path.join(base, 'Parent report.xlsx'))
valid_accounts = ['111006', '120030', '121015', '188600', '433002', '534018']

from app.ic_processor import match_journal_to_icm
lookup_filtered = {k: v for k, v in parent_lookup.items() if k[2] in valid_accounts}
updates = match_journal_to_icm(data_rows, lookup_filtered)

print(f"Updates keys generated:")
for e, p, a in updates:
    if '117100' in e or '007009' in e or '117100' in p or '007009' in p:
        print(f"  ({e}, {p}, {a})")

missing_pairs = set()
for (ent, icp, acct) in updates:
    if acct in valid_accounts and (ent, icp) not in existing_pairs:
        missing_pairs.add((ent, icp))

print(f"\nMissing pairs identified:")
for e, p in missing_pairs:
    if '117100' in e or '007009' in e or '117100' in p or '007009' in p:
        print(f"  ({e}, {p})")
