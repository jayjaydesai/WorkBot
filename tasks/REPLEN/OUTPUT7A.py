import pandas as pd
import os

# Define file paths
output6a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT6A.xlsx"
output7a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT7A.xlsx"

# Load the OUTPUT6A file
output6a_df = pd.read_excel(output6a_path)

# Normalize column names
output6a_df.columns = output6a_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['ITEM NUMBER', 'POSTED QUANTITY', 'STOCK FOR REPLEN']
for col in required_columns:
    if col not in output6a_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT6A.xlsx")

# Step 2: Calculate the TOTAL POSTED QTY for each ITEM NUMBER
output6a_df['TOTAL POSTED QTY'] = output6a_df.groupby('ITEM NUMBER')['POSTED QUANTITY'].transform('sum')

# Step 3: Add the USE LOCATION STATUS column
def determine_location_status(stock_for_replen):
    return "USE" if stock_for_replen > 0 else "NO USE"

output6a_df['USE LOCATION STATUS'] = output6a_df['STOCK FOR REPLEN'].apply(determine_location_status)

# Save the final dataframe as OUTPUT7A
output6a_df.to_excel(output7a_path, index=False)
print(f"OUTPUT7A file saved at: {output7a_path}")

