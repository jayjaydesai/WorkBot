import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")  # Use Azure variable if available
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output12_file = os.path.join(OUTPUT_PATH, "OUTPUT12.csv")
output13_file = os.path.join(OUTPUT_PATH, "OUTPUT13.csv")

# Ensure the source file exists
if not os.path.exists(output12_file):
    raise FileNotFoundError(f"ERROR: OUTPUT12.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output12_file)

# Add calculated column "final/actual/stock/balance"
df["final/actual/stock/balance"] = df["actual/stock"] - df["sales-back/order"]

# Save the result as OUTPUT13.csv
df.to_csv(output13_file, index=False)

print(f"SUCCESS: OUTPUT13.csv generated at {output13_file}")
