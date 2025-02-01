import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output20_file = os.path.join(OUTPUT_PATH, "OUTPUT20.csv")
output21_file = os.path.join(OUTPUT_PATH, "OUTPUT21.csv")

# Ensure the source file exists
if not os.path.exists(output20_file):
    raise FileNotFoundError(f"ERROR: OUTPUT20.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output20_file)

# Ensure required columns exist
required_columns = ["release", "note", "customer/total/bo/qty", "actual/stock"]
missing_columns = [col for col in required_columns if col not in df.columns]
if missing_columns:
    raise KeyError(f"ERROR: Missing required columns in OUTPUT20.csv: {missing_columns}")

# Convert "release" column to string to avoid NaN issues
df["release"] = df["release"].fillna("").astype(str).str.strip()
df["note"] = df["note"].fillna("").astype(str).str.strip()

# Update "release" column based on "note" column conditions
df.loc[df["note"] == "release after", "release"] = "0"
df.loc[df["note"] == "back/order/qty", "release"] = df["customer/total/bo/qty"]
df.loc[df["note"] == "actual-stock", "release"] = df["actual/stock"]

# Save the result as OUTPUT21.csv
df.to_csv(output21_file, index=False)

print(f"SUCCESS: OUTPUT21.csv generated at {output21_file}")
