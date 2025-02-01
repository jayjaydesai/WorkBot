import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output17_file = os.path.join(OUTPUT_PATH, "OUTPUT17.csv")
output18_file = os.path.join(OUTPUT_PATH, "OUTPUT18.csv")

# Ensure the source file exists
if not os.path.exists(output17_file):
    raise FileNotFoundError(f"ERROR: OUTPUT17.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output17_file)

# Ensure necessary columns exist
required_columns = ["note", "balance/actual/stock", "actual/stock"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Required column '{col}' not found in OUTPUT17.csv")

# Ensure "note" column exists and replace NaN with blank
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "balance/actual/stock" to numeric (handle errors)
df["balance/actual/stock"] = pd.to_numeric(df["balance/actual/stock"], errors="coerce")

# Apply filter conditions:
# 1. "note" is blank
# 2. "balance/actual/stock" <= 0
mask = (df["note"] == "") & (df["balance/actual/stock"] <= 0)

# Update "note" column with the remark "actual-stock"
df.loc[mask, "note"] = "actual-stock"

# Save the result as OUTPUT18.csv
df.to_csv(output18_file, index=False)

print(f"SUCCESS: OUTPUT18.csv generated at {output18_file}")
