"""Debug script to find the 'list index out of range' error."""
import sys
import os
import traceback

sys.path.insert(0, os.path.dirname(__file__))

# Set up logging to see all messages
import logging
logging.basicConfig(level=logging.DEBUG, format='%(levelname)s: %(message)s')

from app.ic_processor import process_icm_report

icm_path = r"g:\FCCS\New folder (3)\Intercompany Report_1150_Intercompany Report.xlsx"
parent_journal = r"g:\FCCS\New folder (3)\Journal Report (5).xlsx"
contribution_journal = r"g:\FCCS\New folder (3)\Journal Report (6).xlsx"
plug_journal = r"g:\FCCS\Update files\Journal Report (4).xlsx"
report_inputs = r"g:\FCCS\AI\report_inputs.xlsx"

journal_paths = {
    "parent_journal": parent_journal,
    "contribution_journal": contribution_journal,
    "plugaccount_journal": plug_journal,
}

output_path = r"g:\FCCS\backend\debug_output.xlsx"

try:
    result = process_icm_report(
        icm_path=icm_path,
        journal_paths=journal_paths,
        output_path=output_path,
        report_inputs_path=report_inputs,
    )
    print(f"SUCCESS: Output written to {result}")
except Exception as e:
    print(f"ERROR: {e}")
    traceback.print_exc()
