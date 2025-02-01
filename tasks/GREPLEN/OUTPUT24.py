import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output23_file = os.path.join(OUTPUT_PATH, "OUTPUT23.csv")
output24_file = os.path.join(OUTPUT_PATH, "OUTPUT24.csv")

# Ensure the source file exists
if not os.path.exists(output23_file):
    raise FileNotFoundError(f"ERROR: OUTPUT23.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output23_file)

# Add the "REASON" column based on conditions in the "NOTE" column
def determine_reason(note):
    """Determine the reason based on the NOTE value."""
    if note == "done":
        return "Actual Stock is 0 or less"
    elif note == "release after":
        return "ETA is within 35 days to release maximum"
    elif note == "back/order/qty":
        return "Actual Stock is more than total BO qty"
    elif note == "actual-stock":
        return "Actual Stock is lower but closer to Customer BO to release"
    else:
        return "No specific reason"

df["REASON"] = df["NOTE"].apply(determine_reason)

# Save the updated DataFrame as OUTPUT24.csv
df.to_csv(output24_file, index=False)

print(f"SUCCESS: OUTPUT24.csv generated at {output24_file}")
