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
output20_file = os.path.join(OUTPUT_PATH, "OUTPUT20.csv")
output21_file = os.path.join(OUTPUT_PATH, "OUTPUT21.csv")

# Ensure the source file exists
if not os.path.exists(output20_file):
    raise FileNotFoundError(f"ERROR: OUTPUT20.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output20_file, dtype=str)

# Convert column names to lowercase for consistency
df.columns = [col.lower().strip() for col in df.columns]

# Ensure required columns exist
required_columns = ["release", "note", "customer/total/bo/qty", "actual/stock"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise KeyError(f"ERROR: Missing required columns in OUTPUT20.csv: {missing_columns}")

# Handle NaN values in "release" and "note" columns
df["release"] = df["release"].fillna("").astype(str).str.strip()
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Convert numeric columns safely
df["customer/total/bo/qty"] = pd.to_numeric(df["customer/total/bo/qty"], errors="coerce").fillna(0)
df["actual/stock"] = pd.to_numeric(df["actual/stock"], errors="coerce").fillna(0)

# Update "release" column based on "note" column conditions
df.loc[df["note"] == "release after", "release"] = "0"
df.loc[df["note"] == "back/order/qty", "release"] = df["customer/total/bo/qty"].astype(str)
df.loc[df["note"] == "actual-stock", "release"] = df["actual/stock"].astype(str)

# Save the result as OUTPUT21.csv
df.to_csv(output21_file, index=False)

print(f"SUCCESS: OUTPUT21.csv generated at {output21_file}")
