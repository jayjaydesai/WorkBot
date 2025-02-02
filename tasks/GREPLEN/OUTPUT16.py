import os
import pandas as pd
from datetime import datetime

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output15_file = os.path.join(OUTPUT_PATH, "OUTPUT15.csv")
output16_file = os.path.join(OUTPUT_PATH, "OUTPUT16.csv")

# Ensure the source file exists
if not os.path.exists(output15_file):
    raise FileNotFoundError(f"ERROR: OUTPUT15.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output15_file)

# Convert column names to lowercase for case-insensitive handling
df.columns = [col.lower().strip() for col in df.columns]

# Ensure required columns exist
required_columns = ["eta", "final/po/qty", "sales-back/order"]
for col in required_columns:
    if col not in df.columns:
        raise KeyError(f"ERROR: Missing required column '{col}' in OUTPUT15.csv. Available columns: {list(df.columns)}")

# Convert "eta" column to datetime format, assuming "dd-mm-yyyy"
df["eta"] = pd.to_datetime(df["eta"], format="%d-%m-%Y", errors="coerce")

# Get today's date
today = datetime.today()

# Calculate "number/of/days/eta" (remaining days until ETA)
df["number/of/days/eta"] = (df["eta"] - today).dt.days

# Convert necessary columns to numeric before performing calculations
df["final/po/qty"] = pd.to_numeric(df["final/po/qty"], errors="coerce").fillna(0)
df["sales-back/order"] = pd.to_numeric(df["sales-back/order"], errors="coerce").fillna(0)

# Calculate "eta/qty/differencetotal"
df["eta/qty/differencetotal"] = df["final/po/qty"] - df["sales-back/order"]

# Save the result as OUTPUT16.csv
df.to_csv(output16_file, index=False)

print(f"SUCCESS: OUTPUT16.csv generated at {output16_file}")
