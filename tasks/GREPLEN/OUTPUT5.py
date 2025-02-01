import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output4_file = os.path.join(OUTPUT_PATH, "OUTPUT4.csv")
output5_file = os.path.join(OUTPUT_PATH, "OUTPUT5.csv")

# Ensure OUTPUT4.csv exists
if not os.path.exists(output4_file):
    raise FileNotFoundError(f"OUTPUT4.csv not found in {OUTPUT_PATH}")

# Load OUTPUT4.csv
df = pd.read_csv(output4_file, dtype=str)

# Convert columns to appropriate data types
df["purchase/order/quantity"] = pd.to_numeric(df["purchase/order/quantity"], errors="coerce").fillna(0)
df["eta"] = pd.to_datetime(df["eta"], format="%d-%m-%Y", errors="coerce")

# Calculate "total/number/of/po" for each "part/number"
df["total/number/of/po"] = df.groupby("part/number")["part/number"].transform("count")

# Sort by "part/number" and "eta" to determine sequence
df = df.sort_values(by=["part/number", "eta"])

# Assign sequential "eta/status" based on sorted "eta" within each "part/number"
df["eta/status"] = df.groupby("part/number").cumcount() + 1

# Calculate "final/po/qty" as the sum of "purchase/order/quantity" for the same "eta" within the same "part/number"
df["final/po/qty"] = df.groupby(["part/number", "eta"])["purchase/order/quantity"].transform("sum")

# Convert ETA column back to "dd-mm-yyyy" format
df["eta"] = df["eta"].dt.strftime("%d-%m-%Y")

# Save processed file
df.to_csv(output5_file, index=False)

print(f"SUCCESS: OUTPUT5.csv generated at {output5_file}")
