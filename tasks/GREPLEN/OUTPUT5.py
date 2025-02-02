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
print(f"🔍 DEBUG: Looking for OUTPUT4.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output4_file = os.path.join(OUTPUT_PATH, "OUTPUT4.csv")
output5_file = os.path.join(OUTPUT_PATH, "OUTPUT5.csv")

try:
    # Ensure OUTPUT4.csv exists
    if not os.path.exists(output4_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT4.csv not found in {OUTPUT_PATH}")

    # Load OUTPUT4.csv with UTF-8 encoding
    df = pd.read_csv(output4_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure required columns exist
    required_columns = ["part/number", "purchase/order/quantity", "eta"]
    for col in required_columns:
        if col not in df.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT4.csv. Available columns: {list(df.columns)}")

    # Convert "purchase/order/quantity" to numeric (handle errors & fill invalid values with 0)
    df["purchase/order/quantity"] = pd.to_numeric(df["purchase/order/quantity"], errors="coerce").fillna(0)

    # Convert "eta" column to datetime (handle errors)
    df["eta"] = pd.to_datetime(df["eta"], format="%d-%m-%Y", errors="coerce")

    # Ensure "eta" column is not empty before sorting
    if df["eta"].isna().all():
        raise ValueError("❌ ERROR: All values in 'eta' column are invalid or missing.")

    # Calculate "total/number/of/po" for each "part/number"
    df["total/number/of/po"] = df.groupby("part/number")["part/number"].transform("count")

    # Sort by "part/number" and "eta" to determine sequence
    df = df.sort_values(by=["part/number", "eta"])

    # Assign sequential "eta/status" based on sorted "eta" within each "part/number"
    df["eta/status"] = df.groupby("part/number").cumcount() + 1

    # Calculate "final/po/qty" as the sum of "purchase/order/quantity" for the same "eta" within the same "part/number"
    df["final/po/qty"] = df.groupby(["part/number", "eta"])["purchase/order/quantity"].transform("sum")

    # Convert "eta" column back to "dd-mm-yyyy" format
    df["eta"] = df["eta"].dt.strftime("%d-%m-%Y")

    # Save processed file with UTF-8 encoding
    df.to_csv(output5_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT5.csv generated at {output5_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT4.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
