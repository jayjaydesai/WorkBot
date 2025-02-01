import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output14_file = os.path.join(OUTPUT_PATH, "OUTPUT14.csv")
output15_file = os.path.join(OUTPUT_PATH, "OUTPUT15.csv")

# Ensure the source file exists
if not os.path.exists(output14_file):
    raise FileNotFoundError(f"ERROR: OUTPUT14.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output14_file)

# Replace NaN or empty values in "note" column with an empty string for filtering
df["note"] = df["note"].fillna("")

# Filter where "note" is blank and "final/actual/stock/balance" is >= 0
condition = (df["note"] == "") & (df["final/actual/stock/balance"] >= 0)

# Assign "back/order/qty" values to "note" column where condition is met
df.loc[condition, "note"] = "back/order/qty"

# Save the result as OUTPUT15.csv
df.to_csv(output15_file, index=False)

print(f"SUCCESS: OUTPUT15.csv generated at {output15_file}")
