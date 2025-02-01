import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output13_file = os.path.join(OUTPUT_PATH, "OUTPUT13.csv")
output14_file = os.path.join(OUTPUT_PATH, "OUTPUT14.csv")

# Ensure the source file exists
if not os.path.exists(output13_file):
    raise FileNotFoundError(f"ERROR: OUTPUT13.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output13_file)

# Add calculated column "availability/ratio"
df["availability/ratio"] = (df["actual/stock"] / df["customer/total/bo/qty"]) * 100

# Handle cases where "customer/total/bo/qty" is 0 to avoid division by zero
df["availability/ratio"] = df["availability/ratio"].fillna(0)  # Replace NaN values with 0
df["availability/ratio"] = df["availability/ratio"].round(2)  # Round percentage to 2 decimal places

# Save the result as OUTPUT14.csv
df.to_csv(output14_file, index=False)

print(f"SUCCESS: OUTPUT14.csv generated at {output14_file}")
