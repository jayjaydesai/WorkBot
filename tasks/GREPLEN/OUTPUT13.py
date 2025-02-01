import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

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
