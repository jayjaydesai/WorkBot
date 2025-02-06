import os
import pandas as pd

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output23_file = os.path.join(OUTPUT_PATH, "OUTPUT23.csv")
output24_file = os.path.join(OUTPUT_PATH, "OUTPUT24.csv")

# Ensure the source file exists
if not os.path.exists(output23_file):
    raise FileNotFoundError(f"ERROR: OUTPUT23.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output23_file, dtype=str)

# Ensure "NOTE" column exists
if "NOTE" not in df.columns:
    raise KeyError("ERROR: 'NOTE' column not found in OUTPUT23.csv")

# Define a function to determine the "REASON" column based on "NOTE"
def determine_reason(note):
    """Determine the reason based on the NOTE value."""
    note = note.strip().lower() if pd.notna(note) else ""
    
    if note == "done":
        return "Actual Stock is 0 or less"
    elif note == "release after":
        return "ETA is within 7 days to release maximum"
    elif note == "back/order/qty":
        return "Actual Stock is more than total BO qty"
    elif note == "actual-stock":
        return "Actual Stock is lower but closer to Customer BO to release"
    else:
        return "No specific reason"

# Apply the function to create the "REASON" column
df["REASON"] = df["NOTE"].apply(determine_reason)

# Save the updated DataFrame as OUTPUT24.csv
df.to_csv(output24_file, index=False)

print(f"SUCCESS: OUTPUT24.csv generated at {output24_file}")
