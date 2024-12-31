from tasks.OUTPUT4 import process_bulk_file
from tasks.OUTPUT5 import process_output4
from tasks.OUTPUT6 import process_output5
from tasks.OUTPUT7 import process_output6
from tasks.OUTPUT8 import process_output7
from tasks.OUTPUT9 import process_output8
import os

def run_all_outputs_for_all_aisles(bulk_file_path, aisle_ids=None):
    """
    Run the complete workflow for specified aisles or all aisles if none are specified.

    Args:
        bulk_file_path (str): Path to the uploaded BULK.xlsx file.
        aisle_ids (list or None): List of aisle identifiers (e.g., ['A01', 'A02', ...]).
                                  If None, process all aisles.

    Returns:
        dict: Mapping of aisle IDs to their final processed output paths.
    """
    results = {}

    try:
        # Use environment variable or default path for locations
        LOCATIONS_FOLDER = os.getenv("LOCATIONS_FOLDER", "locations")  # Default to "locations"

        # Default to all aisles (A01 to A18) if none are specified
        if aisle_ids is None:
            aisle_ids = [f"A{str(i).zfill(2)}" for i in range(1, 19)]  # Example: A01 to A18

        for aisle_id in aisle_ids:
            print(f"Processing aisle {aisle_id}...")

            # Check if the aisle file exists
            aisle_file_path = os.path.join(LOCATIONS_FOLDER, f"{aisle_id}.xlsx")
            if not os.path.exists(aisle_file_path):
                print(f"Skipping aisle {aisle_id}: File not found in {LOCATIONS_FOLDER}.")
                continue  # Skip this aisle

            # Process each step in the workflow
            try:
                output4_path = process_bulk_file(bulk_file_path, aisle_id)
                output5_path = process_output4(aisle_id)
                output6_path = process_output5(aisle_id)
                output7_path = process_output6(aisle_id)
                output8_path = process_output7(aisle_id)
                final_output_path = process_output8(aisle_id)

                # Store the final output path
                results[aisle_id] = final_output_path
                print(f"Completed processing for aisle {aisle_id}. Final output: {final_output_path}")

            except Exception as e:
                print(f"Error processing aisle {aisle_id}: {e}")

        print("All aisles processed successfully.")
        return results

    except Exception as e:
        print(f"An error occurred during the workflow: {e}")
        raise Exception(f"Workflow failed: {e}")
