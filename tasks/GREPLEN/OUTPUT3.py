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
print(f"🔍 DEBUG: Looking for PO.csv at: {UPLOAD_PATH}")

# Ensure required directories exist
os.makedirs(UPLOAD_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
po_file = os.path.join(UPLOAD_PATH, "PO.csv")
output3_file = os.path.join(OUTPUT_PATH, "OUTPUT3.csv")

try:
    # Ensure PO.csv exists
    if not os.path.exists(po_file):
        raise FileNotFoundError(f"❌ ERROR: PO.csv not found in {UPLOAD_PATH}")

    # Load PO.csv with UTF-8 encoding
    df = pd.read_csv(po_file, header=None, dtype=str, encoding="utf-8")

    # Define new column headers for the first 11 columns
    new_headers = [
        "site1", "part/number", "site2", "document/type", "purchase/order/number",
        "indication", "supplier", "first/qty", "purchase/order/quantity", "created/date", "eta"
    ]

    # Ensure PO.csv has at least 11 columns; add empty ones if required
    while len(df.columns) < len(new_headers):
        df[len(df.columns)] = None  # Add empty columns

    # Assign new column headers
    df.columns = new_headers[:len(df.columns)]  # Assign available headers

    # Convert "created/date" and "eta" columns to datetime format (dd-mm-yyyy)
    for date_col in ["created/date", "eta"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.strftime("%d-%m-%Y")

    # Save processed file with UTF-8 encoding
    df.to_csv(output3_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT3.csv generated at {output3_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: PO.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
