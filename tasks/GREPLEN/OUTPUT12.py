import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output11_file = os.path.join(OUTPUT_PATH, "OUTPUT11.csv")
output12_file = os.path.join(OUTPUT_PATH, "OUTPUT12.csv")

# Ensure the source file exists
if not os.path.exists(output11_file):
    raise FileNotFoundError(f"ERROR: OUTPUT11.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output11_file)

# Add the calculated column "note" based on conditions
df["note"] = df.apply(
    lambda row: "back/order/qty" if row["difference"] == 0 and row["balance/actual/stock"] >= 0 
    else ("done" if row["release"] == 0 else ""),
    axis=1
)

# Save the result as OUTPUT12.csv
df.to_csv(output12_file, index=False)

print(f"SUCCESS: OUTPUT12.csv generated at {output12_file}")
