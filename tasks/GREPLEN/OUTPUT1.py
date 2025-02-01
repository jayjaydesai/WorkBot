import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
UPLOAD_PATH = os.path.join(BASE_DIR, "uploads", "GREPLEN")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure output directory exists
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
master_file = os.path.join(UPLOAD_PATH, "MASTER.csv")
output1_file = os.path.join(OUTPUT_PATH, "OUTPUT1.csv")

# Ensure MASTER.csv exists
if not os.path.exists(master_file):
    raise FileNotFoundError(f"MASTER.csv not found in {UPLOAD_PATH}")

# Load MASTER.csv
df = pd.read_csv(master_file)

# Process column names: lowercase & replace spaces with "/"
df.columns = [col.strip().lower().replace(" ", "/") for col in df.columns]

# Save processed file
df.to_csv(output1_file, index=False)

print(f"SUCCESS: OUTPUT1.csv generated at {output1_file}")
