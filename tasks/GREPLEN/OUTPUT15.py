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
output14_file = os.path.join(OUTPUT_PATH, "OUTPUT14.csv")
output15_file = os.path.join(OUTPUT_PATH, "OUTPUT15.csv")

# Ensure the source file exists
if not os.path.exists(output14_file):
    raise FileNotFoundError(f"ERROR: OUTPUT14.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output14_file)

# Convert column names to lowercase for case-insensitive handling
df.columns = [col.lower().strip() for col in df.columns]

# Ensure required columns exist
required_columns = ["note", "final/actual/stock/balance"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Missing required column '{col}' in OUTPUT14.csv. Available columns: {list(df.columns)}")

# Replace NaN or empty values in "note" column with an empty string for filtering
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert "final/actual/stock/balance" to numeric before filtering
df["final/actual/stock/balance"] = pd.to_numeric(df["final/actual/stock/balance"], errors="coerce").fillna(0)

# Apply condition where "note" is blank and "final/actual/stock/balance" is >= 0
condition = (df["note"] == "") & (df["final/actual/stock/balance"] >= 0)

# Assign "back/order/qty" to "note" column where condition is met
df.loc[condition, "note"] = "back/order/qty"

# Save the result as OUTPUT15.csv
df.to_csv(output15_file, index=False)

print(f"SUCCESS: OUTPUT15.csv generated at {output15_file}")
