import os
import pandas as pd
from pathlib import Path

# Function to get the latest BULK.xlsx file from the uploads folder
def get_latest_bulk_file(upload_folder):
    files = [f for f in Path(upload_folder).glob("*.xlsx")]
    if not files:
        raise FileNotFoundError("No .xlsx files found in the uploads folder.")
    latest_file = max(files, key=lambda x: x.stat().st_mtime)  # Get the most recently modified file
    return latest_file

def process_files(base_path):
    # Define folder structure dynamically
    locations_folder = Path(base_path) / "locations"
    uploads_folder = Path(base_path) / "uploads"
    output_folder = Path(base_path) / "output"

    # Ensure output folder exists
    output_folder.mkdir(parents=True, exist_ok=True)

    # Paths for input files
    a16_path = locations_folder / "A16.xlsx"
    bulk_path = get_latest_bulk_file(uploads_folder)

    # Load Excel files
    a16_df = pd.read_excel(a16_path)
    bulk_df = pd.read_excel(bulk_path)

    # Merge based on the "Location" column
    merged_df = pd.merge(a16_df, bulk_df[['Location', 'Licence plate']], on='Location', how='left')

    # Save the resulting dataframe to the output path
    output_file = output_folder / "OUTPUT10.xlsx"
    merged_df.to_excel(output_file, index=False)

    print(f"Output saved to: {output_file}")

# Example of using the script dynamically
if __name__ == "__main__":
    # Base path should be the root of your project; dynamically set for Azure
    base_path = Path(__file__).parent  # Adjust this to the root directory
    process_files(base_path)
