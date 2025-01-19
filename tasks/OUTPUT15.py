import pandas as pd
from pathlib import Path

def process_duplicates():
    # Define the correct input file path
    input_file = Path(r"C:\Users\jaydi\OneDrive - Comline\CAPLOCATION\Deployment\bulk_report_webapp\output\OUTPUT11.xlsx")
    print(f"Looking for file at: {input_file}")
    print(f"Checking file existence: {input_file.exists()}")
    print(f"Absolute path being checked: {input_file.resolve()}")

    if not input_file.exists():
        raise FileNotFoundError(f"Source file {input_file} not found.")

    # Define the output file path (OUTPUT15.xlsx)
    output_file = Path(r"C:\Users\jaydi\OneDrive - Comline\CAPLOCATION\Deployment\bulk_report_webapp\output\OUTPUT15.xlsx")

    # Read the Excel file
    df = pd.read_excel(input_file)

    # Check if "Licence plate" column exists
    if "Licence plate" not in df.columns:
        raise ValueError("The 'Licence plate' column is missing in the source file.")

    # Filter rows with duplicate Licence plate values
    duplicate_licence_plates = df[df.duplicated(subset=["Licence plate"], keep=False)]

    # Reset the index for the resulting DataFrame
    duplicate_licence_plates.reset_index(drop=True, inplace=True)

    # Save the filtered data to OUTPUT15.xlsx
    duplicate_licence_plates.to_excel(output_file, index=False)

    print(f"Filtered rows with duplicate Licence plates have been saved to {output_file}")

if __name__ == "__main__":
    process_duplicates()


