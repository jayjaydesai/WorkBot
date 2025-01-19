import pandas as pd
import os

# Define file paths
output1_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT1.xlsx"
coverage_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\2-SOURCE_FILES\COVERAGE.xlsx"
output2_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT2.xlsx"

# Load the data
output1_df = pd.read_excel(output1_path)
coverage_df = pd.read_excel(coverage_path)

# Normalize column names (strip spaces and convert to lowercase for consistency)
output1_df.columns = output1_df.columns.str.strip().str.lower()
coverage_df.columns = coverage_df.columns.str.strip().str.lower()

# Check if 'item number' exists in both files
if 'item number' not in output1_df.columns or 'item number' not in coverage_df.columns:
    raise KeyError("'Item number' column is missing in one of the files.")

# Ensure case-insensitive matching by converting 'item number' to lowercase
output1_df['item number'] = output1_df['item number'].astype(str).str.lower()
coverage_df['item number'] = coverage_df['item number'].astype(str).str.lower()

# Select only the required columns from COVERAGE
required_columns = ['item number', 'description', 'sub range', 'pallet qty', 'brand', 'min', 'coverage']
coverage_subset = coverage_df[required_columns]

# Merge OUTPUT1 and COVERAGE based on 'item number'
merged_df = pd.merge(output1_df, coverage_subset, on='item number', how='left')

# Convert all column headings to uppercase
merged_df.columns = merged_df.columns.str.upper()

# Convert specific text columns to uppercase
columns_to_uppercase = ['ITEM NUMBER', 'WAREHOUSE', 'LOCATION', 'STOCK STATUS', 'LICENCE PLATE']
for col in columns_to_uppercase:
    if col in merged_df.columns:
        merged_df[col] = merged_df[col].astype(str).str.upper()

# Replace blank or missing values in 'SUB RANGE' with 'NEW'
if 'SUB RANGE' in merged_df.columns:
    merged_df['SUB RANGE'] = merged_df['SUB RANGE'].fillna('NEW')

# Save the final merged file
if not os.path.exists(os.path.dirname(output2_path)):
    os.makedirs(os.path.dirname(output2_path))  # Create the output folder if it doesn't exist

merged_df.to_excel(output2_path, index=False)
print(f"Merged file saved at: {output2_path}")
