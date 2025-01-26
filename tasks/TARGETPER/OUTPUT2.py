import os
import pandas as pd

# Define the base paths dynamically
base_path = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
uploads_path = os.path.join(base_path, "uploads", "TARGETPER")
output_path = os.path.join(base_path, "output", "TARGETPER")

# File paths
output1_file = os.path.join(output_path, "OUTPUT1.csv")
sales_modified_file = os.path.join(output_path, "SALES_MODIFIED.csv")
output2_file = os.path.join(output_path, "OUTPUT2.csv")

# Ensure the input files exist
if not os.path.exists(output1_file):
    raise FileNotFoundError(f"OUTPUT1.csv not found in {output_path}")
if not os.path.exists(sales_modified_file):
    raise FileNotFoundError(f"SALES_MODIFIED.csv not found in {output_path}")

# Load the input files
output1_df = pd.read_csv(output1_file)
sales_df = pd.read_csv(sales_modified_file)

# Convert date columns to datetime and handle errors gracefully
sales_df['invoicedate'] = pd.to_datetime(sales_df['invoicedate'], errors='coerce', dayfirst=True)
output1_df['targetdate'] = pd.to_datetime(output1_df['targetdate'], errors='coerce', dayfirst=True)

# Ensure invalid dates are handled (replace NaT with None or unknown)
sales_df['invoicedate'] = sales_df['invoicedate'].dt.strftime('%d-%m-%Y').fillna('unknown')
output1_df['targetdate'] = output1_df['targetdate'].dt.strftime('%d-%m-%Y').fillna('unknown')

# Convert date ranges for filtering sales
sales_df['invoicedate'] = pd.to_datetime(sales_df['invoicedate'], format='%d-%m-%Y', errors='coerce')

# Compute 2023-customer-sales and 2024-customer-sales
def calculate_yearly_sales(output_df, sales_df, year, target_column):
    """Calculate yearly sales for a given year and update the target column in OUTPUT1."""
    start_date = f"01-01-{year}"
    end_date = f"31-12-{year}"
    mask = (sales_df['invoicedate'] >= pd.to_datetime(start_date)) & (sales_df['invoicedate'] <= pd.to_datetime(end_date))
    yearly_sales = sales_df[mask].groupby('standardcode')['qty'].sum()
    output_df[target_column] = output_df['standardcode'].map(yearly_sales).fillna(0).astype(int)

calculate_yearly_sales(output1_df, sales_df, 2023, '2023-customer-sales')
calculate_yearly_sales(output1_df, sales_df, 2024, '2024-customer-sales')

# Compute number-of-customer-target
output1_df['number-of-customer-target'] = output1_df.groupby('standardcode')['standardcode'].transform('count')

# Compute number-of-part-number-target
output1_df['number-of-part-number-target'] = output1_df.groupby('partnumber')['partnumber'].transform('count')

# Save the final DataFrame as OUTPUT2.csv
output1_df.to_csv(output2_file, index=False)

print(f"OUTPUT2.csv generated successfully at: {output2_file}")
