import subprocess
from pathlib import Path

# Define the path relative to the current file's directory
scripts_directory = Path(__file__).parent / "tasks"

# List of scripts to run in sequence
scripts = ["OUTPUT11.py", "OUTPUT12.py", "OUTPUT13.py"]

def run_script(script_name):
    script_path = scripts_directory / script_name
    try:
        print(f"Running {script_name}...")
        # Execute the script
        subprocess.run(["python", str(script_path)], check=True)
        print(f"{script_name} executed successfully.\n")
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while running {script_name}: {e}\n")

if __name__ == "__main__":
    for script in scripts:
        run_script(script)

