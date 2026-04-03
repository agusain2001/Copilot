"""Reproduce the 'list index out of range' error from report sequence 26."""
import sys, os, traceback
sys.path.insert(0, os.path.dirname(__file__))

from app.ic_processor import process_icm_report

icm_path = r"g:\FCCS\backend\uploads\reports\26\inputs\Intercompany Report_1150_Intercompany Report.xlsx"
journal_paths = {
    "parent_journal":       r"g:\FCCS\backend\uploads\reports\26\inputs\Journal Report (5).xlsx",
    "contribution_journal": r"g:\FCCS\backend\uploads\reports\26\inputs\Journal Report (6).xlsx",
    "plugaccount_journal":  r"g:\FCCS\backend\uploads\reports\26\inputs\Journal Report (4).xlsx",
}
report_inputs_path = r"g:\FCCS\backend\uploads\reports\26\inputs\report_inputs.xlsx"
output_path = r"g:\FCCS\backend\uploads\reports\26\outputs\ICM_Output_26.xlsx"

try:
    process_icm_report(icm_path, journal_paths, output_path, report_inputs_path)
    print("SUCCESS: Output written to", output_path)
except Exception as e:
    print(f"ERROR: {e}")
    traceback.print_exc()
