import os
import pandas as pd
from pathlib import Path


def create_output1(upload_folder, output_folder):
    """
    Create OUTPUT1.xlsx by merging BULK.xlsx, AS01011.xlsx, and COVERAGE.xlsx.
    Retains first-letter capitalization in column headers without unnecessary formatting.

    Args:
        upload_folder (str): Path to the folder containing input files.
        output_folder (str): Path to save the output file.
    """
    try:
        # Define required and optional files
        required_files = {
            "bulk_file": "BULK.xlsx",
            "as_file": "AS01011.xlsx",
            "coverage_file": "COVERAGE.xlsx",
        }

        # Resolve full paths for required files
        input_files = {key: Path(upload_folder) / filename for key, filename in required_files.items()}

        # Validate required files exist
        missing_files = [filename for key, filename in required_files.items() if not input_files[key].exists()]
        if missing_files:
            raise FileNotFoundError(f"Missing required files: {', '.join(missing_files)}. Please upload these files.")

        # Resolve output file path
        output_file = Path(output_folder) / "OUTPUT1.xlsx"
        os.makedirs(output_folder, exist_ok=True)  # Ensure the output folder exists

        # Load files and process data
        print("Loading BULK.xlsx...")
        bulk_df = pd.read_excel(input_files["bulk_file"])
        bulk_df.columns = bulk_df.columns.str.lower().str.strip()

        print("Loading AS01011.xlsx...")
        as_df = pd.read_excel(input_files["as_file"], usecols=["Item number", "Available physical"])
        as_df.columns = as_df.columns.str.lower().str.strip()
        as_df["available physical"] = as_df["available physical"].fillna(0)

        print("Loading COVERAGE.xlsx...")
        coverage_df = pd.read_excel(input_files["coverage_file"])
        coverage_df.columns = coverage_df.columns.str.lower().str.strip()

        # Merge BULK.xlsx with COVERAGE.xlsx on "item number"
        print("Merging BULK.xlsx and COVERAGE.xlsx...")
        merged_df = pd.merge(bulk_df, coverage_df, on="item number", how="left")

        # Merge the result with AS01011.xlsx on "item number"
        print("Merging with AS01011.xlsx...")
        final_df = pd.merge(merged_df, as_df, on="item number", how="left")

        # Capitalize the first letter of each word in column headers
        final_df.columns = final_df.columns.str.title()

        # Save the final DataFrame to OUTPUT1.xlsx
        print(f"Saving OUTPUT1.xlsx to: {output_file}")
        final_df.to_excel(output_file, index=False)

        print(f"OUTPUT1.xlsx created successfully at {output_file}")
        return output_file

    except FileNotFoundError as e:
        print(f"File Not Found Error: {str(e)}")
        raise
    except Exception as e:
        print(f"Error occurred while processing: {e}")
        raise


if __name__ == "__main__":
    # Dynamically resolve paths based on the script's environment
    base_path = Path(__file__).resolve().parent.parent.parent
    upload_folder = base_path / "uploads" / "REPLEN"
    output_folder = base_path / "output" / "REPLEN"

    # Debugging: Print resolved paths
    print(f"Resolved upload folder: {upload_folder}")
    print(f"Resolved output folder: {output_folder}")

    # Check if upload folder exists
    if not upload_folder.exists():
        print(f"Upload folder does not exist: {upload_folder}")
        raise FileNotFoundError(f"Upload folder not found: {upload_folder}")

    # List files in the upload folder
    print(f"Files in upload folder: {list(upload_folder.glob('*'))}")

    # Run the function
    create_output1(upload_folder, output_folder)
