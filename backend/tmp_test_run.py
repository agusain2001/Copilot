"""Quick end-to-end test using the correct uploaded filenames."""
import sys, logging
logging.basicConfig(level=logging.INFO, format="%(message)s")
sys.path.insert(0, r"g:\FCCS\backend\app")

from ic_processor import process_icm_report

BASE = r"g:\FCCS\backend\uploads\reports\9\inputs"

result = process_icm_report(
    icm_path          = BASE + r"\Intercompany Balances IC Matching Report (1).xlsx",
    journal_paths     = {
        "parent_journal":       BASE + r"\Journal Report (1).xlsx",
        "contribution_journal": BASE + r"\Journal Report (2).xlsx",
        "plugaccount_journal":  BASE + r"\Journal Report (4).xlsx",
    },
    output_path       = r"g:\FCCS\backend\uploads\reports\9\outputs\ICM_Output_TEST.xlsx",
    report_inputs_path= BASE + r"\report_inputs.xlsx",
)
print("Done:", result)
