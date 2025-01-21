import os
import pandas as pd
from pathlib import Path


def create_output3(output_folder):
    """
    Process OUTPUT2.xlsx to add calculated columns and save as OUTPUT3.xlsx.

    Args:
        output_folder (str): Path to save the output file.
    """
    try:
        # Resolve paths dynamically
        output_folder = Path(output_folder).resolve()

        print(f"Resolved output folder: {output_folder}")

        # Ensure output folder exists
        if not output_folder.exists():
            raise FileNotFoundError(f"Output folder does not exist: {output_folder}")

        # File paths
        output2_file = output_folder / "OUTPUT2.xlsx"
        output3_file = output_folder / "OUTPUT3.xlsx"

        # Check if OUTPUT2.xlsx exists
        if not output2_file.exists():
            raise FileNotFoundError(f"Required file not found: {output2_file}")

        # Load OUTPUT2.xlsx
        print("Loading OUTPUT2.xlsx...")
        output2_df = pd.read_excel(output2_file)

        # Normalize column names for consistency
        output2_df.columns = output2_df.columns.str.lower().str.strip()

        # Add calculated columns
        print("Adding calculated columns...")
        if "rmin" not in output2_df.columns or "available physical" not in output2_df.columns:
            raise ValueError("Required columns 'rmin' or 'available physical' not found in OUTPUT2.xlsx.")

        output2_df["rmin - available physical"] = output2_df["rmin"] - output2_df["available physical"]
        output2_df["replen status"] = output2_df["rmin - available physical"].apply(
            lambda x: "Replen Required" if x > 0 else "Replen not Required"
        )

        # Capitalize the first letter of every word in column headers
        output2_df.columns = output2_df.columns.str.title()

        # Save to OUTPUT3.xlsx
        print(f"Saving OUTPUT3.xlsx to: {output3_file}")
        output2_df.to_excel(output3_file, index=False)
        print(f"OUTPUT3.xlsx created successfully at {output3_file}")

        return output3_file

    except FileNotFoundError as e:
        print(f"File Not Found Error: {str(e)}")
        raise
    except ValueError as e:
        print(f"Value Error: {str(e)}")
        raise
    except Exception as e:
        print(f"Error occurred while processing: {e}")
        raise


if __name__ == "__main__":
    # Dynamically determine paths
    base_path = Path(__file__).resolve().parents[2]  # Go up two levels from this script's location
    output_folder = base_path / "output" / "REPLEN"

    # Run the function
    create_output3(output_folder)
