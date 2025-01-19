import pandas as pd
import os

# Define file paths
output10_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT10.xlsx"
output11_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT11.xlsx"

# Load the OUTPUT10 file
output10_df = pd.read_excel(output10_path)

# Normalize column names
output10_df.columns = output10_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['ITEM NUMBER', 'DIFF', 'STOCK FOR REPLEN', 'POSTED QUANTITY']
for col in required_columns:
    if col not in output10_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT10.xlsx")

# Step 2: Create the "STOCK FOR REPLEN2" column
output10_df['STOCK FOR REPLEN2'] = output10_df['STOCK FOR REPLEN'].apply(
    lambda x: 0 if x == "DECISION" else x
)

# Step 3: Calculate the "BALANCE STOCK FOR REPLEN"
# Group by ITEM NUMBER and sum STOCK FOR REPLEN2
stock_for_replen_sum = output10_df.groupby('ITEM NUMBER')['STOCK FOR REPLEN2'].transform('sum')
output10_df['BALANCE STOCK FOR REPLEN'] = output10_df['DIFF'] - stock_for_replen_sum

# Step 4: Calculate the "TOTAL POSTED QTY2"
# Group by ITEM NUMBER and sum POSTED QUANTITY
output10_df['TOTAL POSTED QTY2'] = output10_df.groupby('ITEM NUMBER')['POSTED QUANTITY'].transform('sum')

# Save the final dataframe as OUTPUT11
output10_df.to_excel(output11_path, index=False)
print(f"OUTPUT11 file saved at: {output11_path}")
