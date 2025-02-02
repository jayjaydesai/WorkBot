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
print(f"🔍 DEBUG: Looking for OUTPUT6.csv and OUTPUT7.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output6_file = os.path.join(OUTPUT_PATH, "OUTPUT6.csv")
output7_file = os.path.join(OUTPUT_PATH, "OUTPUT7.csv")
output8_file = os.path.join(OUTPUT_PATH, "OUTPUT8.csv")

try:
    # Ensure both source files exist
    if not os.path.exists(output6_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT6.csv not found in {OUTPUT_PATH}")
    if not os.path.exists(output7_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT7.csv not found in {OUTPUT_PATH}")

    # Load source files with UTF-8 encoding
    df6 = pd.read_csv(output6_file, dtype=str, encoding="utf-8")
    df7 = pd.read_csv(output7_file, dtype=str, encoding="utf-8")

    # Convert column names to lowercase to avoid case-sensitivity issues
    df6.columns = [col.lower().strip() for col in df6.columns]
    df7.columns = [col.lower().strip() for col in df7.columns]

    # Ensure required columns exist in OUTPUT6.csv
    required_columns_df6 = ["index/part/number"]
    for col in required_columns_df6:
        if col not in df6.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT6.csv. Available columns: {list(df6.columns)}")

    # Ensure required columns exist in OUTPUT7.csv
    required_columns_df7 = [
        "index/part/number", "supplier", "purchase/order/quantity", "eta", 
        "total/po/qty", "total/number/of/po", "final/po/qty", "eta/status"
    ]
    for col in required_columns_df7:
        if col not in df7.columns:
            raise KeyError(f"❌ ERROR: Column '{col}' not found in OUTPUT7.csv. Available columns: {list(df7.columns)}")

    # Convert "index/part/number" to string before merging
    df6["index/part/number"] = df6["index/part/number"].astype(str)
    df7["index/part/number"] = df7["index/part/number"].astype(str)

    # Perform the merge (left join)
    df8 = pd.merge(df6, df7[required_columns_df7], on="index/part/number", how="left")

    # Save the merged file with UTF-8 encoding
    df8.to_csv(output8_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT8.csv generated at {output8_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: One of the input files is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
