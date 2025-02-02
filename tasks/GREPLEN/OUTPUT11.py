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
print(f"🔍 DEBUG: Looking for OUTPUT10.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output10_file = os.path.join(OUTPUT_PATH, "OUTPUT10.csv")
output11_file = os.path.join(OUTPUT_PATH, "OUTPUT11.csv")

try:
    # Ensure the source file exists
    if not os.path.exists(output10_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT10.csv not found in {OUTPUT_PATH}")

    # Load the source file with UTF-8 encoding
    df = pd.read_csv(output10_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure required columns exist
    required_columns = ["customer/total/bo/qty", "sales-back/order", "actual/stock"]
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT10.csv. Available columns: {list(df.columns)}")

    # Convert necessary columns to numeric before calculations
    df["customer/total/bo/qty"] = pd.to_numeric(df["customer/total/bo/qty"], errors="coerce").fillna(0)
    df["sales-back/order"] = pd.to_numeric(df["sales-back/order"], errors="coerce").fillna(0)
    df["actual/stock"] = pd.to_numeric(df["actual/stock"], errors="coerce").fillna(0)

    # Add calculated columns
    df["difference"] = df["customer/total/bo/qty"] - df["sales-back/order"]
    df["balance/actual/stock"] = df["actual/stock"] - df["customer/total/bo/qty"]

    # Save the result as OUTPUT11.csv with UTF-8 encoding
    df.to_csv(output11_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT11.csv generated at {output11_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT10.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
