import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output16_file = os.path.join(OUTPUT_PATH, "OUTPUT16.csv")
output17_file = os.path.join(OUTPUT_PATH, "OUTPUT17.csv")

# Ensure the source file exists
if not os.path.exists(output16_file):
    raise FileNotFoundError(f"ERROR: OUTPUT16.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output16_file)

# Ensure necessary columns exist
required_columns = ["note", "number/of/days/eta", "eta/qty/differencetotal"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Required column '{col}' not found in OUTPUT16.csv")

# Ensure "note" column exists and replace NaN with blank
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "number/of/days/eta" and "eta/qty/differencetotal" to numeric (handle errors)
df["number/of/days/eta"] = pd.to_numeric(df["number/of/days/eta"], errors="coerce")
df["eta/qty/differencetotal"] = pd.to_numeric(df["eta/qty/differencetotal"], errors="coerce")

# Apply filter conditions:
# 1. "note" is blank
# 2. "number/of/days/eta" <= 35
# 3. "eta/qty/differencetotal" >= 0
mask = (df["note"] == "") & (df["number/of/days/eta"] <= 35) & (df["eta/qty/differencetotal"] >= 0)

# Update "note" column with "release after"
df.loc[mask, "note"] = "release after"

# Save the result as OUTPUT17.csv
df.to_csv(output17_file, index=False)

import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output16_file = os.path.join(OUTPUT_PATH, "OUTPUT16.csv")
output17_file = os.path.join(OUTPUT_PATH, "OUTPUT17.csv")

# Ensure the source file exists
if not os.path.exists(output16_file):
    raise FileNotFoundError(f"ERROR: OUTPUT16.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output16_file)

# Ensure necessary columns exist
required_columns = ["note", "number/of/days/eta", "eta/qty/differencetotal"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Required column '{col}' not found in OUTPUT16.csv")

# Ensure "note" column exists and replace NaN with blank
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "number/of/days/eta" and "eta/qty/differencetotal" to numeric (handle errors)
df["number/of/days/eta"] = pd.to_numeric(df["number/of/days/eta"], errors="coerce")
df["eta/qty/differencetotal"] = pd.to_numeric(df["eta/qty/differencetotal"], errors="coerce")

# Apply filter conditions:
# 1. "note" is blank
# 2. "number/of/days/eta" <= 35
# 3. "eta/qty/differencetotal" >= 0
mask = (df["note"] == "") & (df["number/of/days/eta"] <= 35) & (df["eta/qty/differencetotal"] >= 0)

# Update "note" column with "release after"
df.loc[mask, "note"] = "release after"

# Save the result as OUTPUT17.csv
df.to_csv(output17_file, index=False)

print(f"SUCCESS: OUTPUT17.csv generated at {output17_file}")
