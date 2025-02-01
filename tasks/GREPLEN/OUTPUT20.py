import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output19_file = os.path.join(OUTPUT_PATH, "OUTPUT19.csv")
output20_file = os.path.join(OUTPUT_PATH, "OUTPUT20.csv")

# Ensure the source file exists
if not os.path.exists(output19_file):
    raise FileNotFoundError(f"ERROR: OUTPUT19.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output19_file)

# Ensure "note" column exists and replace NaN with blank
if "note" not in df.columns:
    raise KeyError("ERROR: 'note' column not found in OUTPUT19.csv")

df["note"] = df["note"].fillna("").astype(str).str.strip()

# Apply filter: Fill blank cells in "note" column with "zero"
df.loc[df["note"] == "", "note"] = "zero"

# Save the result as OUTPUT20.csv
df.to_csv(output20_file, index=False)

print(f"SUCCESS: OUTPUT20.csv generated at {output20_file}")
