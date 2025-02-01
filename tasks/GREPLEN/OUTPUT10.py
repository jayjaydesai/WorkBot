import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output9_file = os.path.join(OUTPUT_PATH, "OUTPUT9.csv")
output10_file = os.path.join(OUTPUT_PATH, "OUTPUT10.csv")

# Ensure the source file exists
if not os.path.exists(output9_file):
    raise FileNotFoundError(f"OUTPUT9.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output9_file)

# Add the calculated column "release"
df["release"] = df["actual/stock"].apply(lambda x: 0 if x <= 0 else "")

# Save the result as OUTPUT10.csv
df.to_csv(output10_file, index=False)

print(f"SUCCESS: OUTPUT10.csv generated at {output10_file}")
