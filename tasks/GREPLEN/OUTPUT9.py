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
print(f"🔍 DEBUG: Looking for OUTPUT8.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output8_file = os.path.join(OUTPUT_PATH, "OUTPUT8.csv")
output9_file = os.path.join(OUTPUT_PATH, "OUTPUT9.csv")

try:
    # Ensure the source file exists
    if not os.path.exists(output8_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT8.csv not found in {OUTPUT_PATH}")

    # Load the source file with UTF-8 encoding
    df = pd.read_csv(output8_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Define columns to remove (converted to lowercase for consistency)
    columns_to_remove = [
        "record/id", "item/description", "availability", "lost/sales", "closed/qty",
        "weighed/item", "group", "qenim", "comment", "expected/delivery", "balance",
        "company/id", "third/code", "branch/id", "employee", "supplier_x", "supplier_y", "eta/status"
    ]

    # Remove the specified columns if they exist
    existing_columns_to_remove = [col for col in columns_to_remove if col in df.columns]
    df = df.drop(columns=existing_columns_to_remove, errors="ignore")

    # Save the result as OUTPUT9.csv with UTF-8 encoding
    df.to_csv(output9_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT9.csv generated at {output9_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT8.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
