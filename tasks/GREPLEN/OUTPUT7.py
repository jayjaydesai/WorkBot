import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output5_file = os.path.join(OUTPUT_PATH, "OUTPUT5.csv")
output7_file = os.path.join(OUTPUT_PATH, "OUTPUT7.csv")

# Ensure OUTPUT5.csv exists
if not os.path.exists(output5_file):
    raise FileNotFoundError(f"OUTPUT5.csv not found in {OUTPUT_PATH}")

# Load OUTPUT5.csv
df = pd.read_csv(output5_file, dtype=str)

# Add "index/part/number" by concatenating "part/number" and "eta/status" without space
df["index/part/number"] = df["part/number"] + df["eta/status"]

# Save processed file
df.to_csv(output7_file, index=False)

print(f"SUCCESS: OUTPUT7.csv generated at {output7_file}")
