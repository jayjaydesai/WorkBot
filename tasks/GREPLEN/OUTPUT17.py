import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output16_file = os.path.join(OUTPUT_PATH, "OUTPUT16.csv")
output17_file = os.path.join(OUTPUT_PATH, "OUTPUT17.csv")

# Ensure the source file exists
if not os.path.exists(output16_file):
    raise FileNotFoundError(f"ERROR: OUTPUT16.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output16_file)

# Convert column names to lowercase for case-insensitive handling
df.columns = [col.lower().strip() for col in df.columns]

# Ensure required columns exist
required_columns = ["note", "number/of/days/eta", "eta/qty/differencetotal"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Missing required column '{col}' in OUTPUT16.csv. Available columns: {list(df.columns)}")

# Ensure "note" column exists and replace NaN with blank
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "number/of/days/eta" and "eta/qty/differencetotal" to numeric (handle errors)
df["number/of/days/eta"] = pd.to_numeric(df["number/of/days/eta"], errors="coerce").fillna(0)
df["eta/qty/differencetotal"] = pd.to_numeric(df["eta/qty/differencetotal"], errors="coerce").fillna(0)

# Apply filter conditions:
# 1. "note" is blank
# 2. "number/of/days/eta" <= 7
# 3. "eta/qty/differencetotal" >= 0
mask = (df["note"] == "") & (df["number/of/days/eta"] <= 7) & (df["eta/qty/differencetotal"] >= 0)

# Update "note" column with "release after"
df.loc[mask, "note"] = "release after"

# Save the result as OUTPUT17.csv
df.to_csv(output17_file, index=False)

print(f"SUCCESS: OUTPUT17.csv generated at {output17_file}")
