import os
import pandas as pd
from pathlib import Path

# Function to get the BULK.xlsx file directly
def get_bulk_file(upload_folder):
    bulk_file = Path(upload_folder) / "BULK.xlsx"  # Directly look for BULK.xlsx
    print(f"Checking for BULK.xlsx at: {bulk_file}")  # Debugging: Print file path
    if not bulk_file.exists():
        raise FileNotFoundError(f"BULK.xlsx not found in the uploads folder: {upload_folder}")
    return bulk_file

def process_files():
    # Set base_path dynamically to point to the `bulk_report_webapp` directory
    base_path = Path(__file__).resolve().parent.parent

    # Define folder structure dynamically
    locations_folder = base_path / "locations"
    uploads_folder = base_path / "uploads"
    output_folder = base_path / "output"

    print(f"Uploads folder: {uploads_folder}")  # Debugging: Ensure correct path
    print(f"Output folder: {output_folder}")   # Debugging: Ensure correct path

    # Ensure output folder exists
    output_folder.mkdir(parents=True, exist_ok=True)

    # Paths for input files
    a16_path = locations_folder / "A16.xlsx"
    bulk_path = get_bulk_file(uploads_folder)  # Directly get BULK.xlsx

    # Load Excel files
    a16_df = pd.read_excel(a16_path)
    bulk_df = pd.read_excel(bulk_path)

    # Merge based on the "Location" column, keeping only the Licence plate from BULK
    merged_df = pd.merge(
        a16_df.drop(columns=['Licence plate']),  # Drop Licence plate column from A16
        bulk_df[['Location', 'Licence plate']],  # Only keep Location and Licence plate from BULK
        on='Location',
        how='left'
    )

    # Filter rows where Licence plate is empty
    filtered_df = merged_df[merged_df['Licence plate'].isna()].copy()

    # Reset Sr.No. column to maintain sequence
    filtered_df['Sr.No.'] = range(1, len(filtered_df) + 1)

    # Save the filtered dataframe to the output path as OUTPUT12.xlsx
    output_file = output_folder / "OUTPUT12.xlsx"
    filtered_df.to_excel(output_file, index=False)

    print(f"Filtered output saved to: {output_file}")

# Example of using the script dynamically
if __name__ == "__main__":
    process_files()
