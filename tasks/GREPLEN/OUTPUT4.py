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
print(f"🔍 DEBUG: Looking for OUTPUT3.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output3_file = os.path.join(OUTPUT_PATH, "OUTPUT3.csv")
output4_file = os.path.join(OUTPUT_PATH, "OUTPUT4.csv")

try:
    # Ensure OUTPUT3.csv exists
    if not os.path.exists(output3_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT3.csv not found in {OUTPUT_PATH}")

    # Load OUTPUT3.csv with UTF-8 encoding
    df = pd.read_csv(output3_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure "purchase/order/quantity" column exists
    if "purchase/order/quantity" not in df.columns:
        raise KeyError(f"❌ ERROR: 'purchase/order/quantity' column not found in OUTPUT3.csv. Available columns: {list(df.columns)}")

    # Ensure "part/number" column exists
    if "part/number" not in df.columns:
        raise KeyError(f"❌ ERROR: 'part/number' column not found in OUTPUT3.csv. Available columns: {list(df.columns)}")

    # Convert "purchase/order/quantity" to numeric (handle errors & fill invalid values with 0)
    df["purchase/order/quantity"] = pd.to_numeric(df["purchase/order/quantity"], errors="coerce").fillna(0)

    # Calculate "total/po/qty" by summing "purchase/order/quantity" for each "part/number"
    df["total/po/qty"] = df.groupby("part/number")["purchase/order/quantity"].transform("sum")

    # Save processed file with UTF-8 encoding
    df.to_csv(output4_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT4.csv generated at {output4_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT3.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
