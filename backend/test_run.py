import sys
from app.ic_processor import process_icm_report

icm = r'G:\FCCS\backend\uploads\reports\15\inputs\Intercompany Balances IC Matching Report (1).xlsx'
inputs = r'g:\FCCS\backend\uploads\reports\15\inputs\report_inputs.xlsx'
journals = {
    "parent_journal": r'G:\FCCS\backend\uploads\reports\15\inputs\Journal Report (1).xlsx',
    "contribution_journal": r'G:\FCCS\backend\uploads\reports\15\inputs\Journal Report (2).xlsx',
    "plugaccount_journal": r'G:\FCCS\backend\uploads\reports\15\inputs\Journal Report (4).xlsx'
}
out = r'G:\FCCS\backend\uploads\reports\15\outputs\ICM_Output_15_test.xlsx'

print("Running processor...")
res = process_icm_report(icm, journals, out, inputs)
print("Saved to:", res)
