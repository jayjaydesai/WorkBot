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
print(f"🔍 DEBUG: Looking for OUTPUT2.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output2_file = os.path.join(OUTPUT_PATH, "OUTPUT2.csv")
output6_file = os.path.join(OUTPUT_PATH, "OUTPUT6.csv")

try:
    # Ensure OUTPUT2.csv exists
    if not os.path.exists(output2_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT2.csv not found in {OUTPUT_PATH}")

    # Load OUTPUT2.csv with UTF-8 encoding
    df = pd.read_csv(output2_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure required columns exist
    required_columns = ["part/number", "backorder", "current/stock", "sales/orders"]
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT2.csv. Available columns: {list(df.columns)}")

    # Convert necessary columns to numeric types
    df["backorder"] = pd.to_numeric(df["backorder"], errors="coerce").fillna(0)
    df["current/stock"] = pd.to_numeric(df["current/stock"], errors="coerce").fillna(0)
    df["sales/orders"] = pd.to_numeric(df["sales/orders"], errors="coerce").fillna(0)

    # Calculate "customer/total/bo/qty" as the sum of "backorder" for the same "part/number"
    df["customer/total/bo/qty"] = df.groupby("part/number")["backorder"].transform("sum")

    # Calculate "actual/stock" as the difference between "current/stock" and "sales/orders"
    df["actual/stock"] = df["current/stock"] - df["sales/orders"]

    # Calculate "total/number/bo" as the count of "part/number" for the same "part/number"
    df["total/number/bo"] = df.groupby("part/number")["part/number"].transform("count")

    # Add "index/part/number" by appending "1" to the value in "part/number"
    df["index/part/number"] = df["part/number"].astype(str) + "1"

    # Save processed file with UTF-8 encoding
    df.to_csv(output6_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT6.csv generated at {output6_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT2.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
