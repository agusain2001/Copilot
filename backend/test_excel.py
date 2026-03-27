import pandas as pd
import sys

def analyze_excel(file_path):
    print(f"\n--- Analyzing {file_path} ---")
    try:
        # Read without headers to see raw structure easily
        df = pd.read_excel(file_path, header=None)
        print(f"Shape: {df.shape}")
        
        # Print top 40 rows to see header structure
        print("First 40 rows overview:")
        for i in range(min(40, len(df))):
            non_nulls = df.iloc[i].dropna().to_list()
            if non_nulls:
                print(f"Row {i+1} has {len(non_nulls)} non-null cells. First few: {non_nulls[:10]}")
            else:
                print(f"Row {i+1} is mostly empty.")
        
    except Exception as e:
        print(f"Error: {e}")

analyze_excel(r'g:\FCCS\backend\uploads\reports\15\outputs\ICM_Output_15.xlsx')
analyze_excel(r'g:\FCCS\backend\uploads\reports\13\outputs\ICM_Output_13.xlsx')
