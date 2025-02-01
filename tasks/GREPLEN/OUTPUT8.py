import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output6_file = os.path.join(OUTPUT_PATH, "OUTPUT6.csv")
output7_file = os.path.join(OUTPUT_PATH, "OUTPUT7.csv")
output8_file = os.path.join(OUTPUT_PATH, "OUTPUT8.csv")

# Ensure both source files exist
if not os.path.exists(output6_file):
    raise FileNotFoundError(f"OUTPUT6.csv not found in {OUTPUT_PATH}")
if not os.path.exists(output7_file):
    raise FileNotFoundError(f"OUTPUT7.csv not found in {OUTPUT_PATH}")

# Load source files
df6 = pd.read_csv(output6_file, dtype=str)
df7 = pd.read_csv(output7_file, dtype=str)

# Define the columns to merge from OUTPUT7.csv
columns_to_merge = [
    "supplier", "purchase/order/quantity", "eta", 
    "total/po/qty", "total/number/of/po", "final/po/qty", "eta/status"
]

# Perform the merge
df8 = pd.merge(df6, df7[["index/part/number"] + columns_to_merge], 
               on="index/part/number", how="left")

# Save the merged file
df8.to_csv(output8_file, index=False)

print(f"SUCCESS: OUTPUT8.csv generated at {output8_file}")
