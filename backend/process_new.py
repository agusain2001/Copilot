"""
Process NEW journals from 'Update files' directory and generate output + full report.
"""
import sys, openpyxl, os
sys.path.insert(0, '.')
from app.ic_processor import (
    process_icm_report, BLK_BASE, BLK_PAR, BLK_CONT,
    NUM_ACCTS, COL_PLUG, COL_FINAL, FIXED_ENT_COLS, FIXED_PAR_COLS
)

BASE = r"G:\FCCS\Update files"
OUTPUT = r"G:\FCCS\Update files\ICM_Output_NEW.xlsx"

journal_paths = {
    "parent_journal":       os.path.join(BASE, "Journal Report.xlsx"),
    "contribution_journal": os.path.join(BASE, "Journal Report (2).xlsx"),
    "plugaccount_journal":  os.path.join(BASE, "Journal Report (4).xlsx"),
}
icm_path = os.path.join(BASE, "Intercompany Balances IC Matching Report (1).xlsx")

# Check if report_inputs.xlsx exists in Update files, else use the one from reports/22
report_inputs_path = os.path.join(BASE, "report_inputs.xlsx")
if not os.path.exists(report_inputs_path):
    report_inputs_path = r"G:\FCCS\AI\report_inputs.xlsx"
    print(f"Using report_inputs from: {report_inputs_path}")

print("Processing NEW ICM report from 'Update files'...")
print(f"  ICM:    {os.path.basename(icm_path)}")
print(f"  J1:     {os.path.basename(journal_paths['parent_journal'])}")
print(f"  J2:     {os.path.basename(journal_paths['contribution_journal'])}")
print(f"  J4:     {os.path.basename(journal_paths['plugaccount_journal'])}")
print()

process_icm_report(icm_path, journal_paths, OUTPUT, report_inputs_path)
print(f"Output saved to: {OUTPUT}")
print()

# --- Read and count all values ---
wb = openpyxl.load_workbook(OUTPUT, data_only=True)
ws = wb.active

blocks = [
    ("BASE INPUT", BLK_BASE),
    ("PARENT JOURNAL (J1)", BLK_PAR),
    ("CONTRIBUTION JOURNAL (J2)", BLK_CONT),
]

grand_totals = {}

for block_name, blk in blocks:
    ent_start = blk[0]
    par_start = blk[2]
    cell_count = 0
    values_detail = []
    
    for rn in range(33, ws.max_row + 1):
        entity = ws.cell(rn, 1).value
        if not entity: break
        partner = ws.cell(rn, 2).value
        
        for i, (code, *_) in enumerate(FIXED_ENT_COLS):
            v = ws.cell(rn, ent_start + i).value
            if v is not None and v != 0:
                cell_count += 1
                values_detail.append((rn, entity, partner, "Entity", code, v))
        for i, (code, *_) in enumerate(FIXED_PAR_COLS):
            v = ws.cell(rn, par_start + i).value
            if v is not None and v != 0:
                cell_count += 1
                values_detail.append((rn, entity, partner, "Partner", code, v))
    
    grand_totals[block_name] = cell_count
    
    if block_name != "BASE INPUT":
        print("=" * 90)
        print(f"  {block_name} - {cell_count} values")
        print("=" * 90)
        for rn, ent, prt, side, code, val in values_detail:
            print(f"  R{rn}: {str(ent)[:40]} / {str(prt)[:40]}")
            print(f"        {side}-side | Account: {code} | Value: {val:>18,.2f}")
        print()

# Plug Account
print("=" * 90)
print("  PLUG ACCOUNT (J4)")
print("=" * 90)
plug_count = 0
for rn in range(33, ws.max_row + 1):
    entity = ws.cell(rn, 1).value
    if not entity: break
    v = ws.cell(rn, COL_PLUG).value
    if v is not None and v != 0:
        plug_count += 1
        partner = ws.cell(rn, 2).value
        print(f"  R{rn}: {str(entity)[:40]} / {str(partner)[:40]}")
        print(f"        Plug Account | Value: {v:>18,.2f}")

grand_totals["PLUG ACCOUNT (J4)"] = plug_count
print()

# Grand Summary
print("=" * 90)
print("  GRAND SUMMARY")
print("=" * 90)
total_journals = 0
for name, count in grand_totals.items():
    print(f"  {name}: {count} values")
    if name != "BASE INPUT":
        total_journals += count
print(f"  {'=' * 50}")
print(f"  TOTAL JOURNAL VALUES (J1+J2+J4): {total_journals}")
