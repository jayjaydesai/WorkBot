import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
UPLOAD_PATH = os.path.join(BASE_DIR, "uploads", "GREPLEN")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
po_file = os.path.join(UPLOAD_PATH, "PO.csv")
output3_file = os.path.join(OUTPUT_PATH, "OUTPUT3.csv")

# Ensure PO.csv exists
if not os.path.exists(po_file):
    raise FileNotFoundError(f"PO.csv not found in {UPLOAD_PATH}")

# Load PO.csv
df = pd.read_csv(po_file, header=None, dtype=str)  # Load as string initially to avoid parsing issues

# Define new column headers for the first 10 columns
new_headers = [
    "site1", "part/number", "site2", "document/type", "purchase/order/number",
    "indication", "supplier", "first/qty", "purchase/order/quantity", "created/date", "eta"
]

# Ensure we have at least 11 columns (if PO.csv has fewer, fill missing columns with default names)
while len(df.columns) < len(new_headers):
    df[len(df.columns)] = None  # Add empty columns if required

# Assign new column headers to the first 11 columns
df.columns = new_headers[:len(df.columns)]  # Assign as many headers as available

# Convert "created/date" and "eta" columns to datetime format (dd-mm-yyyy)
for date_col in ["created/date", "eta"]:
    if date_col in df.columns:
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.strftime("%d-%m-%Y")

# Save processed file
df.to_csv(output3_file, index=False)

print(f"SUCCESS: OUTPUT3.csv generated at {output3_file}")
