import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import logging; logging.basicConfig(level=logging.WARNING)
from app.ic_processor import read_journal_report, match_journal_to_icm, read_icm_data, detect_icm_header_row, read_icm_account_columns
import openpyxl

parent_path = r"g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx"
lookup = read_journal_report(parent_path)
va = {"111006", "120030", "121015", "188600", "433002", "534018"}
lf = {k: v for k, v in lookup.items() if k[2] in va}

icm_path = r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx"
wb = openpyxl.load_workbook(icm_path, data_only=True)
ws = wb.active
hr, ds = detect_icm_header_row(ws)
data_rows = read_icm_data(ws, ds)
updates = match_journal_to_icm(data_rows, lf)

key1 = ("117100", "007009", "534018")
key2 = ("007009", "117100", "534018")
print(f"updates has {key1}: {key1 in updates}, val={updates.get(key1)}")
print(f"updates has {key2}: {key2 in updates}, val={updates.get(key2)}")
print(f"Total updates keys: {len(updates)}")

# Also check what par_cols the output engine uses
hr2, _ = detect_icm_header_row(ws)
ent_cols, par_cols = read_icm_account_columns(ws, hr2)
print(f"\nent_cols codes: {[c[0] for c in ent_cols]}")
print(f"par_cols codes: {[c[0] for c in par_cols]}")
