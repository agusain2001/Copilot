import pandas as pd
import openpyxl

wb = openpyxl.load_workbook(r'g:\FCCS\AI\Intercompany Balances IC Matching Report 1.xlsx', data_only=True)
ws = wb.active

for i in range(25, 32):
    row_vals = [str(ws.cell(row=i, column=j).value) for j in range(1, 20)]
    print(f"Row {i}: {row_vals}")

print("AutoFilter:", ws.auto_filter.ref)
