import pandas as pd
import os

# Define file paths
output11_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT11.xlsx"
output12_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT12.xlsx"

# Load the OUTPUT11 file
output11_df = pd.read_excel(output11_path)

# Normalize column names
output11_df.columns = output11_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = [
    'ITEM NUMBER', 'STOCK FOR REPLEN2', 'TOTAL POSTED QTY2', 'BALANCE STOCK FOR REPLEN'
]
for col in required_columns:
    if col not in output11_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT11.xlsx")

# Step 2: Calculate "AVAILABLE POSTED QTY TO ADD"
# Group by ITEM NUMBER and sum STOCK FOR REPLEN2
stock_for_replen_sum = output11_df.groupby('ITEM NUMBER')['STOCK FOR REPLEN2'].transform('sum')
output11_df['AVAILABLE POSTED QTY TO ADD'] = output11_df['TOTAL POSTED QTY2'] - stock_for_replen_sum

# Step 3: Calculate "FULL ADD STATUS"
output11_df['FULL ADD STATUS'] = output11_df.apply(
    lambda row: "FULL" if row['AVAILABLE POSTED QTY TO ADD'] < row['BALANCE STOCK FOR REPLEN'] else "BALANCE",
    axis=1
)

# Save the final dataframe as OUTPUT12
output11_df.to_excel(output12_path, index=False)
print(f"OUTPUT12 file saved at: {output12_path}")
