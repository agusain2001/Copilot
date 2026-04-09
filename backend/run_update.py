import sys
sys.path.insert(0, r'G:\FCCS\backend')
from app.ic_processor import process_icm_report

icm = r'G:\FCCS\Update files\Intercompany Balances IC Matching Report (1).xlsx'
journals = {
    "parent_journal":       r'G:\FCCS\Update files\Journal Report.xlsx',
    "contribution_journal": r'G:\FCCS\Update files\Journal Report (2).xlsx',
    "plugaccount_journal":  r'G:\FCCS\Update files\Journal Report (4).xlsx'
}
inputs = r'G:\FCCS\backend\uploads\reports\22\inputs\report_inputs.xlsx'
out = r'G:\FCCS\Update files\ICM_Output_NEW_RUN.xlsx'

print("Running ICM processor...")
print(f"  ICM:     {icm}")
print(f"  Parent:  {journals['parent_journal']}")
print(f"  Contrib: {journals['contribution_journal']}")
print(f"  Plug:    {journals['plugaccount_journal']}")
print(f"  Output:  {out}")

res = process_icm_report(icm, journals, out, inputs)
print(f"\nDone! Saved to: {res}")
