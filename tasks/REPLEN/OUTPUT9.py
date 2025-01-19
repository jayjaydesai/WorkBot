import pandas as pd
import os

# Define file paths
output8a_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT8A.xlsx"
output9_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT9.xlsx"

# Load the OUTPUT8A file
output8a_df = pd.read_excel(output8a_path)

# Normalize column names
output8a_df.columns = output8a_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['ITEM NUMBER', 'LOCATION_X']
for col in required_columns:
    if col not in output8a_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT8A.xlsx")

# Step 2: Add the calculated column "NUMBER OF LOCATIONS"
output8a_df['NUMBER OF LOCATIONS'] = (
    output8a_df.groupby('ITEM NUMBER')['LOCATION_X']
    .transform('count')
)

# Save the final dataframe as OUTPUT9
output8a_df.to_excel(output9_path, index=False)
print(f"OUTPUT9 file saved at: {output9_path}")

