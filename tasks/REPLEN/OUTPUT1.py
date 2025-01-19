import pandas as pd
import os

# Define file paths
source_folder = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\1-DAILY_STOCK_REPORT\2-RENAME_REPORTS"
output_folder = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT"

file_bulk = os.path.join(source_folder, "BULK.xlsx")
file_as01011 = os.path.join(source_folder, "AS01011.xlsx")
output_file = os.path.join(output_folder, "OUTPUT1.xlsx")

# Load the data
bulk_df = pd.read_excel(file_bulk)
as01011_df = pd.read_excel(file_as01011)

# Clean column names (strip spaces for safety)
as01011_df.columns = as01011_df.columns.str.strip()
bulk_df.columns = bulk_df.columns.str.strip()

# Select and rename the "Available physical" column
if 'Available physical' in as01011_df.columns:
    as01011_df = as01011_df[['Item number', 'Available physical']]
    as01011_df.rename(columns={'Available physical': 'AS01011-Available physical'}, inplace=True)
else:
    raise KeyError("Column 'Available physical' not found in AS01011.xlsx")

# Merge files on 'Item number'
if 'Item number' in bulk_df.columns and 'Item number' in as01011_df.columns:
    merged_df = pd.merge(bulk_df, as01011_df, on='Item number', how='left')
else:
    raise KeyError("Column 'Item number' not found in one or both files")

# Convert all column headings to uppercase
merged_df.columns = merged_df.columns.str.upper()

# Save the merged file
if not os.path.exists(output_folder):
    os.makedirs(output_folder)  # Create the output folder if it doesn't exist

merged_df.to_excel(output_file, index=False)
print(f"Merged file saved at: {output_file}")
