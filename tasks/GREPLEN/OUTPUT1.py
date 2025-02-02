import os
import pandas as pd
import sys

# Ensure UTF-8 output encoding to avoid UnicodeEncodeError
sys.stdout.reconfigure(encoding='utf-8')

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")  # Use Azure environment variable if available

if not BASE_DIR:  # If not set, find the correct project root
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

UPLOAD_PATH = os.path.join(BASE_DIR, "uploads", "GREPLEN")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Debugging: Print paths to check correctness
print(f"🔍 DEBUG: BASE_DIR is set to: {BASE_DIR}")
print(f"🔍 DEBUG: Looking for MASTER.csv at: {UPLOAD_PATH}")

# Ensure required directories exist
os.makedirs(UPLOAD_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
master_file = os.path.join(UPLOAD_PATH, "MASTER.csv")
output1_file = os.path.join(OUTPUT_PATH, "OUTPUT1.csv")

try:
    # Ensure MASTER.csv exists
    if not os.path.exists(master_file):
        raise FileNotFoundError(f"❌ ERROR: MASTER.csv not found in {UPLOAD_PATH}")

    # Load MASTER.csv with UTF-8 encoding
    df = pd.read_csv(master_file, encoding="utf-8")

    # Process column names: lowercase & replace spaces with "/"
    df.columns = [col.strip().lower().replace(" ", "/") for col in df.columns]

    # Save processed file with UTF-8 encoding
    df.to_csv(output1_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT1.csv generated at {output1_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: MASTER.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
