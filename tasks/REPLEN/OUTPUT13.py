import pandas as pd
import os

# Define file paths
output12_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT12.xlsx"
output13_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT13.xlsx"

# Load the OUTPUT12 file
output12_df = pd.read_excel(output12_path)

# Normalize column names
output12_df.columns = output12_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = [
    'ITEM NUMBER', 'STOCK FOR REPLEN2', 'TOTAL POSTED QTY2',
    'BALANCE STOCK FOR REPLEN', 'AVAILABLE POSTED QTY TO ADD'
]
for col in required_columns:
    if col not in output12_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT12.xlsx")

# Step 2: Calculate "FINAL STATUS TO ADD"
def calculate_final_status(row):
    available_posted_qty_to_add = row['AVAILABLE POSTED QTY TO ADD']
    balance_stock_for_replen = row['BALANCE STOCK FOR REPLEN']

    # Check conditions for "FULL"
    if available_posted_qty_to_add < balance_stock_for_replen:
        return "FULL"
    elif (
        balance_stock_for_replen < available_posted_qty_to_add <= balance_stock_for_replen * 1.10
    ):
        return "FULL"
    else:
        return "BALANCE"

# Apply the logic to calculate "FINAL STATUS TO ADD"
output12_df['FINAL STATUS TO ADD'] = output12_df.apply(calculate_final_status, axis=1)

# Save the final dataframe as OUTPUT13
output12_df.to_excel(output13_path, index=False)
print(f"OUTPUT13 file saved at: {output13_path}")
