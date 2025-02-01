import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", 
                                              "CAPLOCATION", "Deployment", "bulk_report_webapp"))
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Define file paths
output24_file = os.path.join(OUTPUT_PATH, "OUTPUT24.csv")
output25_file = os.path.join(OUTPUT_PATH, "OUTPUT25.csv")

# Ensure the source file exists
if not os.path.exists(output24_file):
    raise FileNotFoundError(f"ERROR: OUTPUT24.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output24_file)

# Update the "NOTE" column based on conditions
def update_note(note):
    """Update the NOTE value based on specific conditions."""
    if note == "done":
        return "No Allocation"
    elif note == "back/order/qty":
        return "Full Allocation"
    elif note == "release after":
        return "Please check as ETA is closer so not allocated"
    elif note == "actual-stock":
        return "Best Possible Allocated"
    else:
        return note  # Keep the original value if no condition is met

df["NOTE"] = df["NOTE"].apply(update_note)

# Save the updated DataFrame as OUTPUT25.csv
df.to_csv(output25_file, index=False)

print(f"SUCCESS: OUTPUT25.csv generated at {output25_file}")
