import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from app.ic_processor import process_icm_report

# We will monkey-patch the logger to print when the specific row happens
import logging
class DebugLogger(logging.Logger):
    def info(self, msg, *args, **kwargs):
        pass
    def warning(self, msg, *args, **kwargs):
        pass

process_icm_report(
    icm_path=r"g:\FCCS\backend\uploads\reports\31\inputs\IC Elimination Report_188800_Intercompany Balances Plug A_c_1156_Intercompany Report 1.xlsx",
    journal_paths={'parent_journal':r'g:\FCCS\backend\uploads\reports\31\inputs\Parent report.xlsx','contribution_journal':r'g:\FCCS\backend\uploads\reports\31\inputs\Contribution report.xlsx','plugaccount_journal':r'g:\FCCS\backend\uploads\reports\31\inputs\Journal Report (4).xlsx'},
    output_path=r"g:\FCCS\backend\uploads\reports\31\outputs\ICM_Output_31_FINAL.xlsx",
    report_inputs_path=r"g:\FCCS\backend\uploads\reports\31\inputs\report Inputs.xlsx",
)
