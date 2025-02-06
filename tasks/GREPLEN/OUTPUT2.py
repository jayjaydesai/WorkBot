import os
import pandas as pd
import sys

# Ensure UTF-8 output encoding to avoid UnicodeEncodeError
sys.stdout.reconfigure(encoding='utf-8')

# Define dynamic paths for local and Azure compatibility
BASE_DIR = os.getenv("BASE_DIR")  # Use Azure variable if available

if not BASE_DIR:  # If not set, find the correct project root
    BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")

# Debugging: Print paths to check correctness
print(f"🔍 DEBUG: BASE_DIR is set to: {BASE_DIR}")
print(f"🔍 DEBUG: Looking for OUTPUT1.csv at: {OUTPUT_PATH}")

# Ensure required directories exist
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Define file paths
output1_file = os.path.join(OUTPUT_PATH, "OUTPUT1.csv")
output2_file = os.path.join(OUTPUT_PATH, "OUTPUT2.csv")

try:
    # Ensure OUTPUT1.csv exists
    if not os.path.exists(output1_file):
        raise FileNotFoundError(f"❌ ERROR: OUTPUT1.csv not found in {OUTPUT_PATH}")

    # Load OUTPUT1.csv with UTF-8 encoding
    df = pd.read_csv(output1_file, encoding="utf-8")

    # Convert column names to lowercase for case-insensitive checks
    df.columns = [col.lower().strip() for col in df.columns]

    # Ensure "date" column exists
    if "date" not in df.columns:
        raise KeyError(f"❌ ERROR: 'date' column not found in OUTPUT1.csv. Available columns: {list(df.columns)}")

    # Convert "date" column to datetime format (dd-mm-yyyy)
    df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime("%d-%m-%Y")

    # Sort by "part/number" and month extracted from the "date" column
    if "part/number" not in df.columns:
        raise KeyError(f"❌ ERROR: 'part/number' column not found in OUTPUT1.csv. Available columns: {list(df.columns)}")

    # Convert "date" to datetime if not already done
    df["date"] = pd.to_datetime(df["date"], format="%d-%m-%Y", errors="coerce")

    # Extract "month" and "year" from "date"
    df["month"] = df["date"].dt.month
    df["year"] = df["date"].dt.year

    # Sort by "part/number", "year", and "month"
    df = df.sort_values(by=["part/number", "year", "month"], ascending=[True, True, True])

    # Drop the temporary "month" and "year" columns after sorting
    df = df.drop(columns=["month", "year"])

    # Save processed file with UTF-8 encoding
    df.to_csv(output2_file, index=False, encoding="utf-8")

    print(f"✅ SUCCESS: OUTPUT2.csv generated at {output2_file}")

except FileNotFoundError as e:
    print(f"❌ ERROR: {e}")
except pd.errors.EmptyDataError:
    print("❌ ERROR: OUTPUT1.csv is empty or corrupted.")
except Exception as e:
    print(f"❌ UNEXPECTED ERROR: {str(e)}")
