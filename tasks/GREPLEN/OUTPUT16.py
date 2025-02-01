import os
import pandas as pd
from datetime import datetime

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output15_file = os.path.join(OUTPUT_PATH, "OUTPUT15.csv")
output16_file = os.path.join(OUTPUT_PATH, "OUTPUT16.csv")

# Ensure the source file exists
if not os.path.exists(output15_file):
    raise FileNotFoundError(f"ERROR: OUTPUT15.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output15_file)

# Convert "eta" column to datetime format (ensure it's in "dd-mm-yyyy" format)
df["eta"] = pd.to_datetime(df["eta"], format="%d-%m-%Y", errors="coerce")

# Get today's date
today = datetime.today()

# Calculate "number/of/days/eta" (remaining days until ETA)
df["number/of/days/eta"] = (df["eta"] - today).dt.days

# Calculate "eta/qty/differencetotal"
df["eta/qty/differencetotal"] = df["final/po/qty"] - df["sales-back/order"]

# Save the result as OUTPUT16.csv
df.to_csv(output16_file, index=False)

print(f"SUCCESS: OUTPUT16.csv generated at {output16_file}")
