import os
import shutil

# Define source and destination folders
final_files_source = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\FINAL FILES TO SEND TO ROBERT"
final_files_destination = r"D:\Comline India\AS-REPLEN-DAILY\FINAL FILES TO SEND TO ROBERT"

output_files_source = r"C:\Users\jayde\OneDrive\AS-REPLEN-DAILY\OUTPUT"
output_files_destination = r"D:\Comline India\AS-REPLEN-DAILY\OUTPUT"

# Function to copy files from source to destination
def copy_files(source_folder, destination_folder):
    try:
        # Ensure the destination folder exists
        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)
        
        # Iterate through all files in the source folder
        for file_name in os.listdir(source_folder):
            source_file = os.path.join(source_folder, file_name)
            destination_file = os.path.join(destination_folder, file_name)
            
            # Check if it is a file (not a directory)
            if os.path.isfile(source_file):
                shutil.copy(source_file, destination_file)
                print(f"Copied: {source_file} -> {destination_file}")
            else:
                print(f"Skipped (not a file): {source_file}")
    except Exception as e:
        print(f"Error occurred while copying files: {e}")

# Copy files from FINAL FILES TO SEND TO ROBERT
print("Copying files from FINAL FILES TO SEND TO ROBERT...")
copy_files(final_files_source, final_files_destination)

# Copy files from OUTPUT
print("\nCopying files from OUTPUT...")
copy_files(output_files_source, output_files_destination)

print("\nFile copying completed successfully.")
