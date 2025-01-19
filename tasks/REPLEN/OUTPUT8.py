import pandas as pd
import os

# Define file paths
output7_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT7.xlsx"
output8_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT8.xlsx"

# Load the OUTPUT7 file
output7_df = pd.read_excel(output7_path)

# Normalize column names
output7_df.columns = output7_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = [
    'ITEM NUMBER', 'POSTED QUANTITY', 'STOCK FOR REPLEN', 
    'TOTAL POSTED QTY', 'SUB RANGE'
]
for col in required_columns:
    if col not in output7_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT7.xlsx")

# Step 2: Create the USE LOCATION STATUS column
def determine_location_status(row):
    if row['STOCK FOR REPLEN'] == "DECISION" or row['STOCK FOR REPLEN'] == 0:
        return "NO USE"
    return "USE"

output7_df['USE LOCATION STATUS'] = output7_df.apply(determine_location_status, axis=1)

# Step 3: Add a calculated column for BALANCE POSTED QTY
output7_df['BALANCE POSTED QTY'] = output7_df['POSTED QUANTITY'] - output7_df['STOCK FOR REPLEN'].replace("DECISION", 0)

# Step 4: Save the final dataframe as OUTPUT8
if not os.path.exists(os.path.dirname(output8_path)):
    os.makedirs(os.path.dirname(output8_path))  # Create the output folder if it doesn't exist

output7_df.to_excel(output8_path, index=False)
print(f"OUTPUT8 file saved at: {output8_path}")

