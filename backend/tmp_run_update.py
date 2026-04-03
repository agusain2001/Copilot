import sys, logging
sys.path.insert(0, r'G:\FCCS\backend')
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

from app.ic_processor import process_icm_report

icm = r'G:\FCCS\Update files\Intercompany Balances IC Matching Report (1).xlsx'
journals = {
    "parent_journal":       r'G:\FCCS\Update files\Journal Report.xlsx',
    "contribution_journal": r'G:\FCCS\Update files\Journal Report (2).xlsx',
    "plugaccount_journal":  r'G:\FCCS\Update files\Journal Report (4).xlsx'
}
report_inputs = r'G:\FCCS\AI\report_inputs.xlsx'
out = r'G:\FCCS\Update files\ICM_Output_ENHANCED_v12.xlsx'

print("Running IC Matching processor...")
print(f"  ICM Source:      {icm}")
print(f"  Parent:          {journals['parent_journal']}")
print(f"  Contribution:    {journals['contribution_journal']}")
print(f"  Plug:            {journals['plugaccount_journal']}")
print(f"  Report Inputs:   {report_inputs}")
print(f"  Output:          {out}")

res = process_icm_report(icm, journals, out, report_inputs_path=report_inputs)
print(f"\nDone! Saved to: {res}")
