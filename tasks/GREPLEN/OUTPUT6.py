import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output2_file = os.path.join(OUTPUT_PATH, "OUTPUT2.csv")
output6_file = os.path.join(OUTPUT_PATH, "OUTPUT6.csv")

# Ensure OUTPUT2.csv exists
if not os.path.exists(output2_file):
    raise FileNotFoundError(f"OUTPUT2.csv not found in {OUTPUT_PATH}")

# Load OUTPUT2.csv
df = pd.read_csv(output2_file, dtype=str)

# Convert necessary columns to numeric types
df["backorder"] = pd.to_numeric(df["backorder"], errors="coerce").fillna(0)
df["current/stock"] = pd.to_numeric(df["current/stock"], errors="coerce").fillna(0)
df["sales/orders"] = pd.to_numeric(df["sales/orders"], errors="coerce").fillna(0)

# Calculate "customer/total/bo/qty" as the sum of "backorder" for the same "part/number"
df["customer/total/bo/qty"] = df.groupby("part/number")["backorder"].transform("sum")

# Calculate "actual/stock" as the difference between "current/stock" and "sales/orders"
df["actual/stock"] = df["current/stock"] - df["sales/orders"]

# Calculate "total/number/bo" as the count of "part/number" for the same "part/number"
df["total/number/bo"] = df.groupby("part/number")["part/number"].transform("count")

# Add "index/part/number" by appending "1" to the value in "part/number"
df["index/part/number"] = df["part/number"] + "1"

# Save processed file
df.to_csv(output6_file, index=False)

print(f"SUCCESS: OUTPUT6.csv generated at {output6_file}")
