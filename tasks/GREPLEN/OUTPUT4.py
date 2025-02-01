import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output3_file = os.path.join(OUTPUT_PATH, "OUTPUT3.csv")
output4_file = os.path.join(OUTPUT_PATH, "OUTPUT4.csv")

# Ensure OUTPUT3.csv exists
if not os.path.exists(output3_file):
    raise FileNotFoundError(f"OUTPUT3.csv not found in {OUTPUT_PATH}")

# Load OUTPUT3.csv
df = pd.read_csv(output3_file, dtype=str)  # Read as string to avoid type issues

# Ensure purchase/order/quantity column is numeric (handle empty or invalid values as 0)
df["purchase/order/quantity"] = pd.to_numeric(df["purchase/order/quantity"], errors="coerce").fillna(0)

# Calculate total/po/qty by summing purchase/order/quantity for the same part/number
df["total/po/qty"] = df.groupby("part/number")["purchase/order/quantity"].transform("sum")

# Save processed file
df.to_csv(output4_file, index=False)

print(f"SUCCESS: OUTPUT4.csv generated at {output4_file}")
