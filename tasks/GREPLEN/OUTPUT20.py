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
output19_file = os.path.join(OUTPUT_PATH, "OUTPUT19.csv")
output20_file = os.path.join(OUTPUT_PATH, "OUTPUT20.csv")

# Ensure the source file exists
if not os.path.exists(output19_file):
    raise FileNotFoundError(f"ERROR: OUTPUT19.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output19_file)

# Convert column names to lowercase for case-insensitive handling
df.columns = [col.lower().strip() for col in df.columns]

# Ensure "note" column exists
if "note" not in df.columns:
    raise KeyError(f"ERROR: 'note' column not found in OUTPUT19.csv. Available columns: {list(df.columns)}")

# Replace NaN values and blank cells in "note" column with "zero"
df["note"] = df["note"].fillna("").astype(str).str.strip()
df.loc[df["note"] == "", "note"] = "zero"

# Save the result as OUTPUT20.csv
df.to_csv(output20_file, index=False)

print(f"SUCCESS: OUTPUT20.csv generated at {output20_file}")
