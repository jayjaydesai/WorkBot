import pandas as pd
import os

# Define file paths
output5_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT5.xlsx"
output6_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT6.xlsx"

# Load the OUTPUT5 file
output5_df = pd.read_excel(output5_path)

# Normalize column names
output5_df.columns = output5_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['POSTED QUANTITY', 'DIFF', 'AS01011-AVAILABLE PHYSICAL']
for col in required_columns:
    if col not in output5_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT5.xlsx")

# Step 2: Add the calculated column "STOCK FOR REPLEN"
def calculate_stock_for_replen(row):
    posted_quantity = row['POSTED QUANTITY']
    diff = row['DIFF']
    available_physical = row['AS01011-AVAILABLE PHYSICAL']
    
    if posted_quantity < 10:
        return posted_quantity
    elif (diff - available_physical) * 1.25 > posted_quantity:
        return posted_quantity
    else:
        return "DECISION"

output5_df['STOCK FOR REPLEN'] = output5_df.apply(calculate_stock_for_replen, axis=1)

# Save the final dataframe as OUTPUT6
output5_df.to_excel(output6_path, index=False)
print(f"OUTPUT6 file saved at: {output6_path}")
