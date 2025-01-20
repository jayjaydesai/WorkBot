import os
import pandas as pd
from pathlib import Path
import sys

def get_bulk_file(upload_folder):
    """Get the BULK.xlsx file from the uploads folder."""
    bulk_file = Path(upload_folder) / "BULK.xlsx"
    print(f"Checking for BULK.xlsx at: {bulk_file}")
    if not bulk_file.exists():
        raise FileNotFoundError(f"BULK.xlsx not found in uploads folder: {bulk_file}")
    return bulk_file

def process_files():
    try:
        # Log environment details
        print(f"Current working directory: {Path.cwd()}")
        print(f"Python version: {sys.version}")
        print(f"Environment variables: {os.environ}")

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
        print(f"A16.xlsx exists: {a16_path.exists()}")
        bulk_path = get_bulk_file(uploads_folder)

        if not a16_path.exists():
            raise FileNotFoundError(f"A16.xlsx not found in locations folder: {a16_path}")

        # Load Excel files
        print("Loading A16.xlsx...")
        try:
            a16_df = pd.read_excel(a16_path)
        except Exception as e:
            raise RuntimeError(f"Failed to load A16.xlsx: {e}")

        print("Loading BULK.xlsx...")
        try:
            bulk_df = pd.read_excel(bulk_path)
        except Exception as e:
            raise RuntimeError(f"Failed to load BULK.xlsx: {e}")

        # Debugging: Print dataframes' shapes and columns
        print(f"A16.xlsx shape: {a16_df.shape}, columns: {list(a16_df.columns)}")
        print(f"BULK.xlsx shape: {bulk_df.shape}, columns: {list(bulk_df.columns)}")

        # Merge dataframes
        print("Merging dataframes...")
        try:
            merged_df = pd.merge(
                a16_df.drop(columns=['Licence plate']),  # Drop Licence plate column from A16
                bulk_df[['Location', 'Licence plate']],  # Only keep Location and Licence plate from BULK
                on='Location',
                how='left'
            )
        except Exception as e:
            raise RuntimeError(f"Failed to merge dataframes: {e}")

        # Save output file
        output_file = output_folder / "OUTPUT11.xlsx"
        print(f"Saving output to: {output_file}")
        try:
            merged_df.to_excel(output_file, index=False)
        except Exception as e:
            raise RuntimeError(f"Failed to save OUTPUT11.xlsx: {e}")

        print("Processing completed successfully.")
    except Exception as e:
        # Log error to a file
        error_log = Path(__file__).parent / "error_log.txt"
        with open(error_log, "w") as log_file:
            log_file.write(f"Error occurred:\n{str(e)}")
        print(f"Error logged in {error_log}")
        raise

if __name__ == "__main__":
    process_files()

