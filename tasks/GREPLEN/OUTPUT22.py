import pandas as pd
import os

# Define dynamic paths
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline",
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output21_file = os.path.join(OUTPUT_PATH, "OUTPUT21.csv")
output22_file = os.path.join(OUTPUT_PATH, "OUTPUT22.csv")

# Load source file
df = pd.read_csv(output21_file)

# Step 1: Handle "actual-stock" in the "note" column and "availability/ratio"
condition = (df["note"] == "actual-stock") & (df["availability/ratio"] <= 40)
df.loc[condition, "release"] = 0  # Set "release" column to 0
df.loc[condition, "note"] = "No Allocation"  # Replace "actual-stock" with "No Allocation"

# Step 2: Remove unwanted columns
columns_to_remove = [
    "total/number/bo", "index/part/number", "purchase/order/quantity", "difference",
    "balance/actual/stock", "final/actual/stock/balance", "availability/ratio", "eta/qty/differencetotal"
]
df = df.drop(columns=columns_to_remove, errors="ignore")

# Step 3: Update column headers
df.columns = [col.replace("/", " ").upper() for col in df.columns]  # Replace "/" with " " and convert to uppercase

# Save the modified file
df.to_csv(output22_file, index=False)
print(f"SUCCESS: OUTPUT22.csv has been generated at {output22_file}")

