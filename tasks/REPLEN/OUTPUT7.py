import pandas as pd
import os

# Define file paths
output6_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT6.xlsx"
output7_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT7.xlsx"

# Load the OUTPUT6 file
output6_df = pd.read_excel(output6_path)

# Normalize column names
output6_df.columns = output6_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['ITEM NUMBER', 'POSTED QUANTITY']
for col in required_columns:
    if col not in output6_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT6.xlsx")

# Step 2: Calculate the TOTAL POSTED QTY for each ITEM NUMBER
output6_df['TOTAL POSTED QTY'] = output6_df.groupby('ITEM NUMBER')['POSTED QUANTITY'].transform('sum')

# Save the final dataframe as OUTPUT7
output6_df.to_excel(output7_path, index=False)
print(f"OUTPUT7 file saved at: {output7_path}")
