import os
import pandas as pd
import sys

# Ensure UTF-8 output encoding to avoid UnicodeEncodeError
sys.stdout.reconfigure(encoding='utf-8')

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")  # Use Azure variable if available
if not BASE_DIR:
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Debugging: Print paths to check correctness
print(f"🔍 DEBUG: BASE_DIR is set to: {BASE_DIR}")
print(f"🔍 DEBUG: Looking for OUTPUT5.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output5_file = os.path.join(OUTPUT_PATH, "OUTPUT5.csv")
output7_file = os.path.join(OUTPUT_PATH, "OUTPUT7.csv")

try:
    # Ensure OUTPUT5.csv exists
    if not os.path.exists(output5_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT5.csv not found in {OUTPUT_PATH}")

    # Load OUTPUT5.csv with UTF-8 encoding
    df = pd.read_csv(output5_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure required columns exist
    required_columns = ["part/number", "eta/status"]
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT5.csv. Available columns: {list(df.columns)}")

    # Convert "eta/status" to string before concatenation
    df["index/part/number"] = df["part/number"].astype(str) + df["eta/status"].astype(str)

    # Save processed file with UTF-8 encoding
    df.to_csv(output7_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT7.csv generated at {output7_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT5.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
