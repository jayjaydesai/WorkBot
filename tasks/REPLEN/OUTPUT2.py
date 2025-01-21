import os
import pandas as pd
from pathlib import Path


def create_output2(upload_folder, output_folder):
    """
    Process OUTPUT1.xlsx and merge with EXPORT.xlsx to create OUTPUT2.xlsx,
    replacing blank cells in "Available Physical" with 0 and capitalizing column headers.

    Args:
        upload_folder (str): Path to the folder containing uploaded files.
        output_folder (str): Path to save the output file.
    """
    try:
        # Resolve paths dynamically
        upload_folder = Path(upload_folder).resolve()
        output_folder = Path(output_folder).resolve()

        print(f"Resolved upload folder: {upload_folder}")
        print(f"Resolved output folder: {output_folder}")

        # Ensure folders exist
        if not upload_folder.exists():
            raise FileNotFoundError(f"Upload folder does not exist: {upload_folder}")
        output_folder.mkdir(parents=True, exist_ok=True)

        # File paths
        output1_file = output_folder / "OUTPUT1.xlsx"
        export_file = upload_folder / "EXPORT.xlsx"
        output2_file = output_folder / "OUTPUT2.xlsx"

        # Check if required files exist
        if not output1_file.exists():
            raise FileNotFoundError(f"Required file not found: {output1_file}")

        # Load EXPORT.xlsx if available, else create an empty DataFrame
        if export_file.exists():
            print("Loading EXPORT.xlsx...")
            export_df = pd.read_excel(export_file, usecols=["Item Number", "Stock"])
            export_df.columns = export_df.columns.str.lower().str.strip()
        else:
            print(f"Warning: {export_file} not found. Proceeding without it.")
            export_df = pd.DataFrame(columns=["item number", "stock"])

        # Load OUTPUT1.xlsx
        print("Loading OUTPUT1.xlsx...")
        output1_df = pd.read_excel(output1_file)
        output1_df.columns = output1_df.columns.str.lower().str.strip()

        # Replace blank cells in "Available Physical" with 0
        if "available physical" in output1_df.columns:
            output1_df["available physical"] = output1_df["available physical"].fillna(0)

        # Merge OUTPUT1.xlsx with EXPORT.xlsx on "Item Number"
        print("Merging OUTPUT1.xlsx with EXPORT.xlsx...")
        final_df = pd.merge(output1_df, export_df, on="item number", how="left")

        # Capitalize the first letter of every word in column headers
        final_df.columns = final_df.columns.str.title()

        # Save to OUTPUT2.xlsx
        print(f"Saving OUTPUT2.xlsx to: {output2_file}")
        final_df.to_excel(output2_file, index=False)

        print(f"OUTPUT2.xlsx created successfully at {output2_file}")

    except FileNotFoundError as e:
        print(f"File Not Found Error: {str(e)}")
        raise
    except Exception as e:
        print(f"Error occurred while processing: {e}")
        raise


if __name__ == "__main__":
    # Dynamically determine paths
    base_path = Path(__file__).resolve().parents[2]
    upload_folder = base_path / "uploads" / "REPLEN"
    output_folder = base_path / "output" / "REPLEN"

    # Run the function
    create_output2(upload_folder, output_folder)
