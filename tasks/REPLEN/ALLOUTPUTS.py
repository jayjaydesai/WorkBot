import subprocess
import os

# Define the directory containing all the scripts
scripts_directory = r"D:\Comline India\REPLENSCRIPTS"

# List all the scripts in the order they should be executed
scripts_to_run = [
    "OUTPUT0.py",
    "OUTPUT0A.py",
    "OUTPUT1.py",
    "OUTPUT2.py",
    "OUTPUT3.py",
    "OUTPUT3A.py",
    "OUTPUT4.py",
    "OUTPUT5.py",
    "OUTPUT6A.py",
    "OUTPUT7A.py",
    "OUTPUT8A.py",
    "OUTPUT9.py",
    "OUTPUT10.py",
    "OUTPUT11.py",
    "OUTPUT12.py",
    "OUTPUT13.py",
    "OUTPUT14.py",
    "OUTPUT15.py",
    "OUTPUT16.py",
    "OUTPUT17.py",
    "OUTPUT18C.py",
    "OUTPUT19.py",
    "OUTPUT20.py",
    "E-MAIL.py"
]

# Execute each script in order
for script in scripts_to_run:
    script_path = os.path.join(scripts_directory, script)
    print(f"Running {script}...")
    try:
        # Execute the script
        result = subprocess.run(["python", script_path], capture_output=True, text=True)
        # Print the output for each script
        if result.returncode == 0:
            print(f"{script} executed successfully.")
            print(result.stdout)
        else:
            print(f"Error while running {script}:")
            print(result.stderr)
    except Exception as e:
        print(f"Failed to execute {script} due to error: {e}")

print("All scripts executed.")
