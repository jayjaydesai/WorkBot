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
output13_file = os.path.join(OUTPUT_PATH, "OUTPUT13.csv")
output14_file = os.path.join(OUTPUT_PATH, "OUTPUT14.csv")

# Ensure the source file exists
if not os.path.exists(output13_file):
    raise FileNotFoundError(f"ERROR: OUTPUT13.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output13_file)

# Convert necessary columns to numeric to avoid calculation errors
df["actual/stock"] = pd.to_numeric(df["actual/stock"], errors="coerce").fillna(0)
df["customer/total/bo/qty"] = pd.to_numeric(df["customer/total/bo/qty"], errors="coerce").fillna(0)

# Calculate "availability/ratio" while avoiding division by zero
df["availability/ratio"] = df.apply(
    lambda row: (row["actual/stock"] / row["customer/total/bo/qty"]) * 100 if row["customer/total/bo/qty"] > 0 else 0,
    axis=1
)

# Round percentage to 2 decimal places
df["availability/ratio"] = df["availability/ratio"].round(2)

# Save the result as OUTPUT14.csv
df.to_csv(output14_file, index=False)

print(f"SUCCESS: OUTPUT14.csv generated at {output14_file}")
