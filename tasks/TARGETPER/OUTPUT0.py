import os
import pandas as pd

# Define the base paths dynamically
base_path = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
uploads_path = os.path.join(base_path, "uploads", "TARGETPER")
output_path = os.path.join(base_path, "output", "TARGETPER")

# Ensure the paths exist
if not os.path.exists(uploads_path):
    raise FileNotFoundError(f"Uploads directory does not exist: {uploads_path}")
if not os.path.exists(output_path):
    os.makedirs(output_path)

# File paths
sales_file = os.path.join(uploads_path, "SALES.csv")
target_file = os.path.join(uploads_path, "TARGET.csv")

# Ensure the files exist
if not os.path.exists(sales_file):
    raise FileNotFoundError(f"SALES.csv not found in {uploads_path}")
if not os.path.exists(target_file):
    raise FileNotFoundError(f"TARGET.csv not found in {uploads_path}")

# Load the CSV files
sales_df = pd.read_csv(sales_file)
target_df = pd.read_csv(target_file)

# General Modifications
def preprocess_columns(df):
    """Standardize column headers."""
    df.columns = [col.strip().lower().replace(" ", "") for col in df.columns]
    return df

sales_df = preprocess_columns(sales_df)
target_df = preprocess_columns(target_df)

# Add standardcode column
sales_df["standardcode"] = sales_df["partnumber"].fillna("unknown") + sales_df["customercode"].fillna("unknown")
target_df["standardcode"] = target_df["partnumber"].fillna("unknown") + target_df["customercode"].fillna("unknown")

# Standardize date formats
def standardize_date(df, date_col):
    """Standardize date formats to dd-mm-yyyy."""
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.strftime("%d-%m-%Y")
    df[date_col] = df[date_col].fillna("unknown")  # Replace invalid dates with "unknown"
    return df

sales_df = standardize_date(sales_df, "invoicedate")
target_df = standardize_date(target_df, "targetdate")

# Change Column Sequence
sales_columns = ["standardcode", "partnumber", "customercode", "productgroup", "invoicedate", "invoicenumber", "qty"]
target_columns = ["standardcode", "partnumber", "customercode", "customername", "targetdate", "status"]

sales_df = sales_df[sales_columns]
target_df = target_df[target_columns]

# Save the modified files
modified_sales_file = os.path.join(output_path, "SALES_MODIFIED.csv")
modified_target_file = os.path.join(output_path, "TARGET_MODIFIED.csv")

sales_df.to_csv(modified_sales_file, index=False)
target_df.to_csv(modified_target_file, index=False)

print(f"Modified SALES file saved to: {modified_sales_file}")
print(f"Modified TARGET file saved to: {modified_target_file}")
