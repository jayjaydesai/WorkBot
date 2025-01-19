import pandas as pd
import os

# Define file paths
output9_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT9.xlsx"
output10_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT\OUTPUT10.xlsx"

# Load the OUTPUT9 file
output9_df = pd.read_excel(output9_path)

# Normalize column names
output9_df.columns = output9_df.columns.str.strip().str.upper()

# Step 1: Check for required column
if 'USE LOCATION STATUS' not in output9_df.columns:
    raise KeyError("'USE LOCATION STATUS' column is missing in OUTPUT9.xlsx")

# Step 2: Filter out rows where "USE LOCATION STATUS" is "NO USE"
output10_df = output9_df[output9_df['USE LOCATION STATUS'] != "NO USE"]

# Save the filtered dataframe as OUTPUT10
output10_df.to_excel(output10_path, index=False)
print(f"OUTPUT10 file saved at: {output10_path}")
