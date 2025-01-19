import os
import shutil

# Define source and destination folders
source_folder = r"C:\Users\jaydi\OneDrive - Comline\Daily_Stock_Report\Rename_Reports"
destination_folder_1 = r"D:\Comline India\AS-REPLEN-DAILY\1-DAILY_STOCK_REPORT\2-RENAME_REPORTS"
destination_folder_2 = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\1-DAILY_STOCK_REPORT\2-RENAME_REPORTS"

# Function to copy files
def copy_files_to_multiple_destinations(source, destinations):
    try:
        for destination in destinations:
            # Ensure the destination folder exists
            if not os.path.exists(destination):
                os.makedirs(destination)
                print(f"Created destination folder: {destination}")
            
            # Iterate through files in the source folder
            for file_name in os.listdir(source):
                source_file = os.path.join(source, file_name)
                destination_file = os.path.join(destination, file_name)
                
                # Copy only files (not directories)
                if os.path.isfile(source_file):
                    shutil.copy(source_file, destination_file)
                    print(f"Copied: {source_file} -> {destination_file}")
                else:
                    print(f"Skipped (not a file): {source_file}")
        
        print("\nAll files copied successfully to all destinations!")
    except Exception as e:
        print(f"Error while copying files: {e}")

# List of destination folders
destination_folders = [destination_folder_1, destination_folder_2]

# Execute the copy function
print("Starting file copy process...")
copy_files_to_multiple_destinations(source_folder, destination_folders)
