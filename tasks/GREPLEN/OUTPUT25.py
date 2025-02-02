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
output24_file = os.path.join(OUTPUT_PATH, "OUTPUT24.csv")
output25_file = os.path.join(OUTPUT_PATH, "OUTPUT25.csv")

# Ensure the source file exists
if not os.path.exists(output24_file):
    raise FileNotFoundError(f"ERROR: OUTPUT24.csv not found in {OUTPUT_PATH}")

# Load the source file
df = pd.read_csv(output24_file, dtype=str)

# Ensure "NOTE" column exists
if "NOTE" not in df.columns:
    raise KeyError("ERROR: 'NOTE' column not found in OUTPUT24.csv")

# Define a function to update "NOTE" column based on specific conditions
def update_note(note):
    """Update the NOTE value based on specific conditions."""
    note = note.strip().lower() if pd.notna(note) else ""
    
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

# Apply the function to modify "NOTE" column
df["NOTE"] = df["NOTE"].apply(update_note)

# Save the updated DataFrame as OUTPUT25.csv
df.to_csv(output25_file, index=False)

print(f"SUCCESS: OUTPUT25.csv generated at {output25_file}")
