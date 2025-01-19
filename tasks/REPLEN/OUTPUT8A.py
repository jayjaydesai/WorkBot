import pandas as pd
import os

# Define file paths
output7a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT7A.xlsx"
output8a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT8A.xlsx"

# Load the OUTPUT7A file
output7a_df = pd.read_excel(output7a_path)

# Normalize column names
output7a_df.columns = output7a_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['STOCK FOR REPLEN', 'POSTED QUANTITY']
for col in required_columns:
    if col not in output7a_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT7A.xlsx")

# Step 2: Remove rows where STOCK FOR REPLEN is 0 (optional; if removal is not needed, comment out this block)
filtered_df = output7a_df[output7a_df['STOCK FOR REPLEN'] != 0].copy()

# Step 3: Add the BALANCE POSTED QTY column
filtered_df['BALANCE POSTED QTY'] = filtered_df['POSTED QUANTITY'] - filtered_df['STOCK FOR REPLEN']

# Save the final dataframe as OUTPUT8A
if not os.path.exists(os.path.dirname(output8a_path)):
    os.makedirs(os.path.dirname(output8a_path))  # Create the output folder if it doesn't exist

filtered_df.to_excel(output8a_path, index=False)
print(f"OUTPUT8A file saved at: {output8a_path}")
