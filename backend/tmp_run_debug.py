"""Run process and capture debug log."""
import sys, os, logging
sys.path.insert(0, os.path.dirname(__file__))
logging.basicConfig(level=logging.WARNING, format="%(message)s",
                    filename="tmp_debug_log.txt", filemode="w")
from app.ic_processor import process_icm_report
process_icm_report(
    icm_path=r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx",
    journal_paths={
        "parent_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx",
        "contribution_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Contribution report.xlsx",
        "plugaccount_journal": r"g:\FCCS\backend\uploads\reports\31\inputs\Journal Report (4).xlsx",
    },
    output_path=r"g:\FCCS\backend\uploads\reports\31\outputs\ICM_Output_31_DEBUG.xlsx",
    report_inputs_path=r"g:\FCCS\backend\uploads\reports\31\inputs\report Inputs.xlsx",
)
print("Done")
