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
print(f"🔍 DEBUG: Looking for OUTPUT9.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output9_file = os.path.join(OUTPUT_PATH, "OUTPUT9.csv")
output10_file = os.path.join(OUTPUT_PATH, "OUTPUT10.csv")

try:
    # Ensure the source file exists
    if not os.path.exists(output9_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT9.csv not found in {OUTPUT_PATH}")

    # Load the source file with UTF-8 encoding
    df = pd.read_csv(output9_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure "actual/stock" column exists
    if "actual/stock" not in df.columns:
        raise KeyError(f"❌ ERROR: Column 'actual/stock' not found in OUTPUT9.csv. Available columns: {list(df.columns)}")

    # Convert "actual/stock" to numeric
    df["actual/stock"] = pd.to_numeric(df["actual/stock"], errors="coerce").fillna(0)

    # Add the calculated column "release"
    df["release"] = df["actual/stock"].apply(lambda x: 0 if x <= 0 else "")

    # Save the result as OUTPUT10.csv with UTF-8 encoding
    df.to_csv(output10_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT10.csv generated at {output10_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT9.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
