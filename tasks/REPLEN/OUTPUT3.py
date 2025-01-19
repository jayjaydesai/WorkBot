import pandas as pd
import os

# Define file paths
output2_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT2.xlsx"
pick_location_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\2-SOURCE_FILES\PICK_LOCATION.xlsx"
as_replen_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\2-SOURCE_FILES\AS_REPLEN.xlsx"
output3_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT3.xlsx"

# Load the data
output2_df = pd.read_excel(output2_path)
pick_location_df = pd.read_excel(pick_location_path)
as_replen_df = pd.read_excel(as_replen_path)

# Normalize column names (strip spaces and convert to lowercase for consistency)
output2_df.columns = output2_df.columns.str.strip().str.lower()
pick_location_df.columns = pick_location_df.columns.str.strip().str.lower()
as_replen_df.columns = as_replen_df.columns.str.strip().str.lower()

# Ensure case-insensitive matching by converting 'item number' to lowercase for matching
output2_df['item number'] = output2_df['item number'].astype(str).str.lower()
pick_location_df['item number'] = pick_location_df['item number'].astype(str).str.lower()
as_replen_df['item number'] = as_replen_df['item number'].astype(str).str.lower()

# Select required columns from PICK_LOCATION and AS_REPLEN
pick_location_subset = pick_location_df[['item number', 'location']]
as_replen_subset = as_replen_df[['item number']].rename(columns={'item number': 'as replen number'})

# Merge OUTPUT2 and PICK_LOCATION based on 'item number'
merged_df = pd.merge(output2_df, pick_location_subset, on='item number', how='left')

# Debugging: Print column names before the next merge
print("Columns in merged_df before merging AS_REPLEN:", merged_df.columns.tolist())
print("Columns in as_replen_subset:", as_replen_subset.columns.tolist())

# Add a new column 'LEVEL' based on the last character of 'LOCATION_X'
location_column = 'location_x' if 'location_x' in merged_df.columns else 'location'
if location_column in merged_df.columns:
    merged_df['level'] = merged_df[location_column].astype(str).str[-1]
else:
    raise KeyError(f"'{location_column}' column is missing in the merged dataframe.")

# Ensure consistent column naming for merge
merged_df.rename(columns={'item number': 'item number'}, inplace=True)

# Merge with AS_REPLEN to add 'AS REPLEN NUMBER'
merged_df = pd.merge(merged_df, as_replen_subset, left_on='item number', right_on='as replen number', how='left')

# Convert all relevant columns to uppercase, including 'ITEM NUMBER' and 'AS REPLEN NUMBER'
columns_to_uppercase = ['item number', 'as replen number', 'location', 'location_x']
for col in columns_to_uppercase:
    if col in merged_df.columns:
        merged_df[col] = merged_df[col].astype(str).str.upper()

# Convert all column headings to uppercase
merged_df.columns = merged_df.columns.str.upper()

# Save the final merged file
if not os.path.exists(os.path.dirname(output3_path)):
    os.makedirs(os.path.dirname(output3_path))  # Create the output folder if it doesn't exist

merged_df.to_excel(output3_path, index=False)
print(f"Merged file saved at: {output3_path}")
