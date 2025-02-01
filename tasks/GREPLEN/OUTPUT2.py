import os
import pandas as pd

# Define dynamic paths for local and Azure
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output1_file = os.path.join(OUTPUT_PATH, "OUTPUT1.csv")
output2_file = os.path.join(OUTPUT_PATH, "OUTPUT2.csv")

# Ensure OUTPUT1.csv exists
if not os.path.exists(output1_file):
    raise FileNotFoundError(f"OUTPUT1.csv not found in {OUTPUT_PATH}")

# Load OUTPUT1.csv
df = pd.read_csv(output1_file)

# Check if "date" column exists
if "date" not in df.columns:
    raise KeyError(f"'date' column not found in OUTPUT1.csv. Available columns: {list(df.columns)}")

# Convert "date" column to datetime format (dd-mm-yyyy)
df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%d-%m-%Y")

# Save processed file
df.to_csv(output2_file, index=False)

print(f"SUCCESS: OUTPUT2.csv generated at {output2_file}")
