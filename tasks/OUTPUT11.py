import os
import pandas as pd
from pathlib import Path

def get_bulk_file(upload_folder):
    """Get the BULK.xlsx file from the uploads folder."""
    bulk_file = Path(upload_folder) / "BULK.xlsx"
    print(f"Checking for BULK.xlsx at: {bulk_file}")
    if not bulk_file.exists():
        raise FileNotFoundError(f"BULK.xlsx not found in uploads folder: {bulk_file}")
    return bulk_file

def process_files():
    try:
        # Set base_path dynamically
        base_path = Path(__file__).resolve().parent.parent

        # Define folder structure
        locations_folder = base_path / "locations"
        uploads_folder = base_path / "uploads"
        output_folder = base_path / "output"

        # Debugging: Print paths
        print(f"Base path: {base_path}")
        print(f"Locations folder: {locations_folder}")
        print(f"Uploads folder: {uploads_folder}")
        print(f"Output folder: {output_folder}")

        # Ensure folders exist
        output_folder.mkdir(parents=True, exist_ok=True)

        # Validate input files
        a16_path = locations_folder / "A16.xlsx"
        if not a16_path.exists():
            raise FileNotFoundError(f"A16.xlsx not found in locations folder: {a16_path}")
        bulk_path = get_bulk_file(uploads_folder)

        # Load Excel files
        print("Loading A16.xlsx...")
        a16_df = pd.read_excel(a16_path)
        print("Loading BULK.xlsx...")
        bulk_df = pd.read_excel(bulk_path)

        # Debugging: Print dataframes' shapes
        print(f"A16.xlsx shape: {a16_df.shape}")
        print(f"BULK.xlsx shape: {bulk_df.shape}")

        # Merge dataframes
        print("Merging dataframes...")
        merged_df = pd.merge(
            a16_df.drop(columns=['Licence plate']),  # Drop Licence plate column from A16
            bulk_df[['Location', 'Licence plate']],  # Only keep Location and Licence plate from BULK
            on='Location',
            how='left'
        )

        # Save output file
        output_file = output_folder / "OUTPUT11.xlsx"
        print(f"Saving output to: {output_file}")
        merged_df.to_excel(output_file, index=False)

        print("Processing completed successfully.")
    except Exception as e:
        print(f"Error occurred: {str(e)}")
        raise

if __name__ == "__main__":
    process_files()
