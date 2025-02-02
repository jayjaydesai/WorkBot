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
print(f"🔍 DEBUG: Looking for OUTPUT11.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output11_file = os.path.join(OUTPUT_PATH, "OUTPUT11.csv")
output12_file = os.path.join(OUTPUT_PATH, "OUTPUT12.csv")

try:
    # Ensure the source file exists
    if not os.path.exists(output11_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT11.csv not found in {OUTPUT_PATH}")

    # Load the source file with UTF-8 encoding
    df = pd.read_csv(output11_file, encoding="utf-8")

    # Add the calculated column "note" based on conditions (Original Logic)
    df["note"] = df.apply(
        lambda row: "back/order/qty" if row["difference"] == 0 and row["balance/actual/stock"] >= 0 
        else ("done" if row["release"] == 0 else ""),
        axis=1
    )

    # Save the result as OUTPUT12.csv with UTF-8 encoding
    df.to_csv(output12_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT12.csv generated at {output12_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT11.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
