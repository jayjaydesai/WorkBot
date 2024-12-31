from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
import os

def process_output8(aisle_id):
    """
    Process the OUTPUT8.xlsx file for the specified aisle by applying A4 printing layout 
    and saving the result as OUTPUT9.

    Args:
        aisle_id (str): Aisle identifier (e.g., 'A09').

    Returns:
        str: Path to the processed OUTPUT9 file for the aisle.
    """
    try:
        # Use environment variables for folder paths
        OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "output")  # Default to "output"

        # File paths
        input_file = os.path.join(OUTPUT_FOLDER, f"OUTPUT8_{aisle_id}.xlsx")  # Input file for the aisle
        output_file = os.path.join(OUTPUT_FOLDER, f"OUTPUT9_{aisle_id}.xlsx")  # Output file for the aisle

        # Load the OUTPUT8 workbook
        workbook = load_workbook(input_file)
        sheet = workbook.active  # Get the active worksheet

        # Set page layout for A4 printing
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4  # Set to A4 size
        sheet.page_setup.orientation = "portrait"  # Set to portrait orientation
        sheet.page_setup.fitToWidth = 1  # Shrink content to fit all columns on one page
        sheet.page_setup.fitToHeight = 1  # Shrink content to fit all rows on one page

        # Set custom margins (1 cm = approximately 0.3937 inches)
        sheet.page_margins = PageMargins(
            left=0.5, right=0.5, top=0.3937, bottom=0.3937, header=0.3, footer=0.3
        )

        # Add page numbers in the footer
        sheet.oddFooter.center.text = "Page &[Page] of &[Pages]"  # Footer text for page numbers

        # Save the updated workbook as OUTPUT9
        workbook.save(output_file)
        print(f"Processed file for {aisle_id} saved as {output_file}")

        return output_file

    except Exception as e:
        print(f"Error processing aisle {aisle_id}: {e}")
        raise Exception(f"Failed to process OUTPUT9 for aisle {aisle_id}: {e}")
