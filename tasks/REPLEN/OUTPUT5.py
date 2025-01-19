import pandas as pd
import os

# Define file paths
output4_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT4.xlsx"
output5_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT5.xlsx"

# Load the OUTPUT4 file
output4_df = pd.read_excel(output4_path)

# Normalize column names
output4_df.columns = output4_df.columns.str.strip().str.upper()

# Step 1: Check for required columns
required_columns = ['REPLEN REQUIRED STATUS', 'POSTED QUANTITY', 'LEVEL', 'LOCATION_Y']
for col in required_columns:
    if col not in output4_df.columns:
        raise KeyError(f"Required column '{col}' is missing in OUTPUT4.xlsx")

# Step 2: Remove rows where REPLEN REQUIRED STATUS is "NO REPLEN REQUIRED"
filtered_df = output4_df[output4_df['REPLEN REQUIRED STATUS'] != "NO REPLEN REQUIRED"]

# Step 3: Remove rows where POSTED QUANTITY is 1 or less
filtered_df = filtered_df[filtered_df['POSTED QUANTITY'] > 1]

# Step 4: Remove rows where LEVEL is "S"
filtered_df = filtered_df[filtered_df['LEVEL'] != "S"]

# Step 5: Remove rows where LOCATION_Y begins with "U" or "u"
filtered_df = filtered_df[~filtered_df['LOCATION_Y'].str.startswith(('U', 'u'), na=False)]

# Save the final dataframe as OUTPUT5
filtered_df.to_excel(output5_path, index=False)
print(f"Final OUTPUT5 file saved at: {output5_path}")
