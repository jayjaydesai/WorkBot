import pandas as pd
import os

# Define file paths
output15_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT15.xlsx"
output16_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT16.xlsx"

# Load the OUTPUT15 file
output15_df = pd.read_excel(output15_path)

# Normalize column names
output15_df.columns = output15_df.columns.str.strip().str.upper()

# Step 1: Check for the required columns
required_columns = ['SUB RANGE', 'FINAL STOCK FOR REPLEN']
missing_columns = [col for col in required_columns if col not in output15_df.columns]
if missing_columns:
    raise KeyError(f"Required columns are missing in OUTPUT15.xlsx: {missing_columns}")

# Step 2: Filter rows where SUB RANGE is FAST or MED
output16_df = output15_df[output15_df['SUB RANGE'].isin(['FAST', 'MED','UP'])]

# Step 3: Remove rows where FINAL STOCK FOR REPLEN is 0
output16_df = output16_df[output16_df['FINAL STOCK FOR REPLEN'] != 0]

# Save the filtered DataFrame as OUTPUT16
try:
    output16_df.to_excel(output16_path, index=False)
    print(f"OUTPUT16 file saved at: {output16_path}")
except Exception as e:
    print(f"An error occurred while saving OUTPUT16: {e}")
