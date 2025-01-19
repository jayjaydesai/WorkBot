import pandas as pd
import os

# Define file paths
output0a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT0A.xlsx"
output16_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT16.xlsx"
output16a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT16A.xlsx"

# Load the OUTPUT0A and OUTPUT16 files
output0a_df = pd.read_excel(output0a_path)
output16_df = pd.read_excel(output16_path)

# Normalize column names
output0a_df.columns = output0a_df.columns.str.strip().str.upper()
output16_df.columns = output16_df.columns.str.strip().str.upper()

# Verify available columns in both DataFrames
print("OUTPUT0A Columns:", output0a_df.columns)
print("OUTPUT16 Columns:", output16_df.columns)

# Check for the required columns
required_columns_output0a = ['LOCATION', 'CATEGORY STATUS', 'STOCK FOR REPLEN']
required_columns_output16 = ['LOCATION_X', 'FINAL STOCK FOR REPLEN']

for col in required_columns_output0a:
    if col not in output0a_df.columns:
        raise KeyError(f"Missing column in OUTPUT0A: {col}")
for col in required_columns_output16:
    if col not in output16_df.columns:
        raise KeyError(f"Missing column in OUTPUT16: {col}")

# Rename 'LOCATION' in OUTPUT0A to 'LOCATION_X' for merging
output0a_df.rename(columns={'LOCATION': 'LOCATION_X'}, inplace=True)

# Merge OUTPUT16 with OUTPUT0A on 'LOCATION_X'
merged_df = pd.merge(output16_df, output0a_df, on='LOCATION_X', how='left')

# Step 1: Update FINAL STOCK FOR REPLEN for EXPORT rows
if 'STOCK FOR REPLEN' in merged_df.columns and 'CATEGORY STATUS' in merged_df.columns:
    merged_df['FINAL STOCK FOR REPLEN'] = merged_df.apply(
        lambda row: row['STOCK FOR REPLEN'] if row['CATEGORY STATUS'] == 'EXPORT' else row['FINAL STOCK FOR REPLEN'],
        axis=1
    )
else:
    raise KeyError("Required columns 'STOCK FOR REPLEN' or 'CATEGORY STATUS' are missing in merged_df")

# Step 2: Add REPLEN STATUS column
if 'CATEGORY STATUS' in merged_df.columns and 'LEVEL' in merged_df.columns:
    def calculate_replen_status(row):
        if row['CATEGORY STATUS'] == 'EXPORT':
            return 'USE' if pd.notna(row['LEVEL']) else 'NO USE'
        return 'NO USE'

    merged_df['REPLEN STATUS'] = merged_df.apply(calculate_replen_status, axis=1)
else:
    raise KeyError("Required columns 'CATEGORY STATUS' or 'LEVEL' are missing in merged_df")

# Save the merged DataFrame as OUTPUT16A
try:
    merged_df.to_excel(output16a_path, index=False)
    print(f"OUTPUT16A file saved at: {output16a_path}")
except Exception as e:
    print(f"An error occurred while saving OUTPUT16A: {e}")
