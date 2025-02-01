import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output8_file = os.path.join(OUTPUT_PATH, "OUTPUT8.csv")
output9_file = os.path.join(OUTPUT_PATH, "OUTPUT9.csv")

# Ensure the source file exists
if not os.path.exists(output8_file):
    raise FileNotFoundError(f"OUTPUT8.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output8_file)

# Columns to remove
columns_to_remove = [
    "record/id", "item/description", "availability", "lost/sales", "closed/qty",
    "weighed/item", "group", "qenim", "comment", "expected/delivery", "balance",
    "company/id", "third/code", "branch/id", "employee", "supplier_x", "supplier_y", "eta/status"
]

# Remove the specified columns
df = df.drop(columns=columns_to_remove, errors="ignore")

# Save the result as OUTPUT9.csv
df.to_csv(output9_file, index=False)

print(f"SUCCESS: OUTPUT9.csv generated at {output9_file}")
