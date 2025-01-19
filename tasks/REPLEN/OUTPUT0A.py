import pandas as pd

# File paths
base_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\1-DAILY_STOCK_REPORT\2-RENAME_REPORTS"
output_path = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT"
exportstock_file = f"{base_path}\EXPORTSTOCK.xlsx"
bulk_file = f"{base_path}\BULK.xlsx"
as01011_file = f"{base_path}\AS01011.xlsx"
output_file = f"{output_path}\OUTPUT0A.xlsx"

# Load Excel files into DataFrames
exportstock_df = pd.read_excel(exportstock_file)
bulk_df = pd.read_excel(bulk_file)
as01011_df = pd.read_excel(as01011_file)

# Standardize column names
exportstock_df.rename(columns=lambda x: x.strip().upper(), inplace=True)
bulk_df.rename(columns=lambda x: x.strip().upper(), inplace=True)
as01011_df.rename(columns=lambda x: x.strip().upper(), inplace=True)

# Merge EXPORTSTOCK with BULK on 'ITEM NUMBER'
merged_df = pd.merge(exportstock_df, bulk_df, on="ITEM NUMBER", how="left")

# Add 'Available physical' column from AS01011 and rename it
as01011_df.rename(columns={"AVAILABLE PHYSICAL": "AS01011-AVAILABLE PHYSICAL"}, inplace=True)
merged_df = pd.merge(merged_df, as01011_df[["ITEM NUMBER", "AS01011-AVAILABLE PHYSICAL"]], on="ITEM NUMBER", how="left")

# Replace blank cells in 'AS01011-AVAILABLE PHYSICAL' with 0
merged_df["AS01011-AVAILABLE PHYSICAL"].fillna(0, inplace=True)

# Save the resulting DataFrame to OUTPUT0A.xlsx in the specified output location
merged_df.to_excel(output_file, index=False)

print(f"OUTPUT0A.xlsx has been created successfully at {output_file}")
