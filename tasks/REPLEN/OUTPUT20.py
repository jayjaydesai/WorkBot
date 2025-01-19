import os
import shutil

# Define source and destination folders
source_folder = r"D:\Comline India\AS-REPLEN-DAILY\FINAL FILES TO SEND TO ROBERT"
destination_folder = r"C:\Users\jaydi\OneDrive - Comline\Daily_Stock_Report\FINAL REPORT TO SEND TO ROBERT"

# Function to copy files
def copy_files(source, destination):
    try:
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
        
        print("\nAll files copied successfully!")
    except Exception as e:
        print(f"Error while copying files: {e}")

# Execute the copy function
print("Starting file copy process...")
copy_files(source_folder, destination_folder)
