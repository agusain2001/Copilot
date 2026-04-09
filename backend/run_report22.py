import sys
sys.path.insert(0, r'G:\FCCS\backend')
from app.ic_processor import process_icm_report

icm = r'G:\FCCS\backend\uploads\reports\22\inputs\Intercompany Balances IC Matching Report (1).xlsx'
inputs = r'G:\FCCS\backend\uploads\reports\22\inputs\report_inputs.xlsx'
journals = {
    "parent_journal":       r'G:\FCCS\backend\uploads\reports\22\inputs\Journal Report (1).xlsx',
    "contribution_journal": r'G:\FCCS\backend\uploads\reports\22\inputs\Journal Report (2).xlsx',
    "plugaccount_journal":  r'G:\FCCS\backend\uploads\reports\22\inputs\Journal Report (4).xlsx'
}
# Use different filename to avoid file lock issue
out = r'G:\FCCS\backend\uploads\reports\22\outputs\ICM_Output_22_v2.xlsx'

print("Running IC Matching processor for Report 22...")
print(f"  Parent: {journals['parent_journal']}")
print(f"  Output: {out}")

res = process_icm_report(icm, journals, out, inputs)
print(f"\nDone! Saved to: {res}")
