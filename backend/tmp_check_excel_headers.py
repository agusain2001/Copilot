import pandas as pd

try:
    df2 = pd.read_excel(r'g:\FCCS\AI\Intercompany Balances IC Matching Report 1.xlsx', header=None)
    # find the first row that is non-null
    for i in range(25, 36):
        print(f"Row {i}:")
        row_vals = df2.iloc[i].tolist()
        print([str(x) for x in row_vals if pd.notna(x)])
except Exception as e:
    print('Error df2:', e)
