import pandas as pd

try:
    df1 = pd.read_excel(r'g:\FCCS\AI\ICM_Output_7.xlsx', header=None)
    df1.head(10).to_csv(r'g:\FCCS\backend\tmp_check_out1.csv', index=False, header=False)
except Exception as e:
    print('Error df1:', e)

try:
    df2 = pd.read_excel(r'g:\FCCS\AI\Intercompany Balances IC Matching Report 1.xlsx', header=None)
    df2.head(10).to_csv(r'g:\FCCS\backend\tmp_check_out2.csv', index=False, header=False)
except Exception as e:
    print('Error df2:', e)
