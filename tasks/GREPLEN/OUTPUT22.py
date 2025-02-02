import pandas as pd
import os

# Define dynamic paths
BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output21_file = os.path.join(OUTPUT_PATH, "OUTPUT21.csv")
output22_file = os.path.join(OUTPUT_PATH, "OUTPUT22.csv")

# Ensure the source file exists
if not os.path.exists(output21_file):
    raise FileNotFoundError(f"ERROR: OUTPUT21.csv not found in {OUTPUT_PATH}")

# Load source file
df = pd.read_csv(output21_file, dtype=str)

# Convert column names to lowercase for consistency
df.columns = [col.lower().strip() for col in df.columns]

# Ensure required columns exist
required_columns = ["note", "availability/ratio", "release"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise KeyError(f"ERROR: Missing required columns in OUTPUT21.csv: {missing_columns}")

# Convert "availability/ratio" to numeric safely
df["availability/ratio"] = pd.to_numeric(df["availability/ratio"], errors="coerce").fillna(0)

# Apply "No Allocation" rule
condition = (df["note"] == "actual-stock") & (df["availability/ratio"] <= 40)
df.loc[condition, "release"] = "0"
df.loc[condition, "note"] = "No Allocation"

# Remove unwanted columns
columns_to_remove = [
    "total/number/bo", "index/part/number", "purchase/order/quantity", "difference",
    "balance/actual/stock", "final/actual/stock/balance", "availability/ratio", "eta/qty/differencetotal"
]
df = df.drop(columns=columns_to_remove, errors="ignore")

# Format column headers: Replace "/" with " " and convert to uppercase
df.columns = [col.replace("/", " ").upper() for col in df.columns]

# Save the modified file
df.to_csv(output22_file, index=False)

print(f"SUCCESS: OUTPUT22.csv has been generated at {output22_file}")
