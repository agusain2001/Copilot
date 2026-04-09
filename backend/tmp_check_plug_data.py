import sys
sys.path.insert(0, r'G:\FCCS\backend')
from app.ic_processor import read_journal_report, parse_report_inputs

journals = {
    "parent": r'G:\FCCS\Update files\Journal Report.xlsx',
    "contrib": r'G:\FCCS\Update files\Journal Report (2).xlsx',
    "plug": r'G:\FCCS\Update files\Journal Report (4).xlsx'
}
report_inputs = r'G:\FCCS\AI\report_inputs.xlsx'
plug_map = parse_report_inputs(report_inputs)
plug_code = plug_map["plug_code"]
print(f"Plug Code is {plug_code}")

for j_name, j_path in journals.items():
    prim, fall = read_journal_report(j_path, None)  # without plug map
    # check how many 188800 records
    count_188 = sum(1 for k in prim.keys() if k[2] == plug_code)
    # check how many elim accounts
    count_elim = sum(1 for k in prim.keys() if k[2] in plug_map["elim_codes"])
    
    print(f"{j_name} journal (no mapping): {count_188} plug, {count_elim} elim")

    # Now with map
    prim_m, fall_m = read_journal_report(j_path, plug_map)
    count_188_m = sum(1 for k in prim_m.keys() if k[2] == plug_code)
    print(f"{j_name} journal (with mapping): {count_188_m} plug")
