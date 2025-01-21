import os
import pandas as pd
from pathlib import Path


def create_output4(output_folder):
    """
    Add a calculated "Difference" column to OUTPUT3.xlsx and save it as OUTPUT4.xlsx.

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
        output3_file = output_folder / "OUTPUT3.xlsx"
        output4_file = output_folder / "OUTPUT4.xlsx"

        # Check if OUTPUT3.xlsx exists
        if not output3_file.exists():
            raise FileNotFoundError(f"Required file not found: {output3_file}")

        # Load OUTPUT3.xlsx into a DataFrame
        print("Loading OUTPUT3.xlsx...")
        df = pd.read_excel(output3_file)

        # Normalize column names for consistency
        df.columns = df.columns.str.lower().str.strip()

        # Add "Difference" column: rcoverage - available physical
        print("Adding 'Difference' column...")
        if "rcoverage" not in df.columns or "available physical" not in df.columns:
            raise ValueError("Required columns 'rcoverage' or 'available physical' not found in OUTPUT3.xlsx.")
        df["difference"] = df["rcoverage"] - df["available physical"]

        # Capitalize the first letter of every word in column headers
        df.columns = df.columns.str.title()

        # Save to OUTPUT4.xlsx
        print(f"Saving OUTPUT4.xlsx to: {output4_file}")
        df.to_excel(output4_file, index=False)
        print(f"OUTPUT4.xlsx created successfully at {output4_file}")

        return output4_file

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
    create_output4(output_folder)
