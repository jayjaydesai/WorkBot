import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output22_file = os.path.join(OUTPUT_PATH, "OUTPUT22.csv")
output23_file = os.path.join(OUTPUT_PATH, "OUTPUT23.csv")

# Ensure the source file exists
if not os.path.exists(output22_file):
    raise FileNotFoundError(f"ERROR: OUTPUT22.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output22_file)

# Desired column sequence
column_order = [
    "REC ID", "DATE", "DESCRIPTION", "DOCUMENT NO", "COMPANY NAME", 
    "PART NUMBER", "UNIT PRICE", "QTY", "PRIORITY CODE", "PRSYPAPPSEC", 
    "CURRENT STOCK", "SALES ORDERS", "BACKORDER", "CUSTOMER TOTAL BO QTY", 
    "SALES-BACK ORDER", "ACTUAL STOCK", "ETA", "FINAL PO QTY", 
    "NUMBER OF DAYS ETA", "RELEASE", "NOTE"
]

# Check if all required columns are present
missing_columns = [col for col in column_order if col not in df.columns]
if missing_columns:
    raise ValueError(f"ERROR: Missing columns in source file: {missing_columns}")

# Reorder the columns
df = df[column_order]

# Save the updated file as OUTPUT23.csv
df.to_csv(output23_file, index=False)

print(f"SUCCESS: OUTPUT23.csv generated at {output23_file}")
