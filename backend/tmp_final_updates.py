"""Reproduce the EXACT flow of process_icm_report and check updates_list_final."""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
import openpyxl, logging
logging.basicConfig(level=logging.WARNING)

from app.ic_processor import *

icm_path = r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx"
journal_paths = {
    "parent_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx",
    "contribution_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Contribution report.xlsx",
    "plugaccount_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Journal Report (4).xlsx",
}
report_inputs_path = r"g:\FCCS\backend\uploads\reports\31\inputs\report Inputs.xlsx"

plug_mapping = parse_report_inputs(report_inputs_path)
valid_accounts = {'111006', '120030', '121015', '188600', '433002', '534018'}

# Re-match step (same as lines 889-904)
journal_order = ["parent_journal", "contribution_journal", "plugaccount_journal"]
updates_list_final = []
for jkey in journal_order:
    jpath = journal_paths.get(jkey)
    if not jpath:
        updates_list_final.append({})
        continue
    jmap = plug_mapping if jkey == "plugaccount_journal" else None
    lookup = read_journal_report(jpath, jmap)
    lookup_filtered = {k: v for k, v in lookup.items() if k[2] in valid_accounts}
    updates = match_journal_to_icm([], lookup_filtered)  # data_rows not used
    updates_list_final.append(updates)

parent_updates = updates_list_final[0]

out = open("tmp_final_updates.txt", "w", encoding="utf-8")
out.write(f"Parent updates count: {len(parent_updates)}\n")
key = ("117100", "007009", "534018")
out.write(f"Key {key} in parent_updates? {key in parent_updates}\n")
if key in parent_updates:
    out.write(f"  Value: {parent_updates[key]}\n")

# Check what write_output receives and does
out.write("\n=== ALL parent updates for 117100 ===\n")
for (e, p, a), v in sorted(parent_updates.items()):
    if e == "117100" or p == "117100":
        out.write(f"  ({e}, {p}, {a}) -> {v}\n")

out.close()
print("Done")
