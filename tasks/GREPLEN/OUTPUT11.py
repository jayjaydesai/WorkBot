import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output10_file = os.path.join(OUTPUT_PATH, "OUTPUT10.csv")
output11_file = os.path.join(OUTPUT_PATH, "OUTPUT11.csv")

# Ensure the source file exists
if not os.path.exists(output10_file):
    raise FileNotFoundError(f"OUTPUT10.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output10_file)

# Add calculated columns
df["difference"] = df["customer/total/bo/qty"] - df["sales-back/order"]
df["balance/actual/stock"] = df["actual/stock"] - df["customer/total/bo/qty"]

# Save the result as OUTPUT11.csv
df.to_csv(output11_file, index=False)

print(f"SUCCESS: OUTPUT11.csv generated at {output11_file}")
