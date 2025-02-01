import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output18_file = os.path.join(OUTPUT_PATH, "OUTPUT18.csv")
output19_file = os.path.join(OUTPUT_PATH, "OUTPUT19.csv")

# Ensure the source file exists
if not os.path.exists(output18_file):
    raise FileNotFoundError(f"ERROR: OUTPUT18.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output18_file)

# Ensure necessary columns exist
required_columns = ["note", "availability/ratio"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Required column '{col}' not found in OUTPUT18.csv")

# Ensure "note" column exists and replace NaN with blank
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "availability/ratio" to numeric (handle errors)
df["availability/ratio"] = pd.to_numeric(df["availability/ratio"], errors="coerce")

# Apply filter conditions:
# 1. "note" is blank
# 2. "availability/ratio" >= 90
mask = (df["note"] == "") & (df["availability/ratio"] >= 90)

# Update "note" column with the remark "back/order/qty"
df.loc[mask, "note"] = "back/order/qty"

# Save the result as OUTPUT19.csv
df.to_csv(output19_file, index=False)

print(f"SUCCESS: OUTPUT19.csv generated at {output19_file}")
