import os
import pandas as pd

# Define dynamic paths
base_path = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
output_path = os.path.join(base_path, "output", "TARGETPER")

# File paths
output2_file = os.path.join(output_path, "OUTPUT2.csv")
sales_modified_file = os.path.join(output_path, "SALES_MODIFIED.csv")
output3_file = os.path.join(output_path, "OUTPUT3.csv")

# Ensure the input files exist
if not os.path.exists(output2_file):
    raise FileNotFoundError(f"OUTPUT2.csv not found in {output_path}")
if not os.path.exists(sales_modified_file):
    raise FileNotFoundError(f"SALES_MODIFIED.csv not found in {output_path}")

# Load the CSV files
output2_df = pd.read_csv(output2_file)
sales_df = pd.read_csv(sales_modified_file)

# Convert date columns to datetime format and handle errors
output2_df['targetdate'] = pd.to_datetime(output2_df['targetdate'], errors='coerce', dayfirst=True)
sales_df['invoicedate'] = pd.to_datetime(sales_df['invoicedate'], errors='coerce', dayfirst=True)

# Drop rows with invalid dates in both dataframes
sales_df = sales_df.dropna(subset=['invoicedate'])
output2_df = output2_df.dropna(subset=['targetdate'])

# Pre-filter SALES_MODIFIED.csv for valid years to reduce processing
sales_df = sales_df[sales_df['invoicedate'] >= output2_df['targetdate'].min()]

# Merge OUTPUT2.csv with SALES_MODIFIED.csv on 'standardcode'
merged_df = pd.merge(
    output2_df,
    sales_df,
    on='standardcode',
    how='left',
    suffixes=('', '_sales')
)

# Filter rows where invoicedate is >= targetdate
merged_df = merged_df[merged_df['invoicedate'] >= merged_df['targetdate']]

# Group by 'standardcode' and 'targetdate' to sum the 'qty'
sales_after_target = merged_df.groupby(['standardcode', 'targetdate'])['qty'].sum().reset_index()

# Merge the aggregated results back into OUTPUT2.csv
output2_df = pd.merge(
    output2_df,
    sales_after_target,
    on=['standardcode', 'targetdate'],
    how='left'
)

# Replace NaN values in 'qty' (now sales-after-target) with 0 and rename column
output2_df['sales-after-target'] = output2_df['qty'].fillna(0).astype(int)
output2_df.drop(columns=['qty'], inplace=True)

# Save the updated file
output2_df.to_csv(output3_file, index=False)

print(f"OUTPUT3.csv generated successfully at: {output3_file}")
