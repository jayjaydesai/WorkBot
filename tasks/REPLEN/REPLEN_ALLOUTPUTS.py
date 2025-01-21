import subprocess
from pathlib import Path

def main():
    # Resolve the base directory dynamically
    base_directory = Path(__file__).resolve().parent

    # Avoid duplication of 'tasks/REPLEN' by directly defining the correct path
    if base_directory.name == "REPLEN":
        scripts_directory = base_directory
    else:
        scripts_directory = base_directory / "tasks" / "REPLEN"

    # Validate the scripts directory
    if not scripts_directory.exists():
        print(f"Scripts directory not found: {scripts_directory}")
        return

    # List all the scripts to be executed in order
    scripts_to_run = [f"OUTPUT{i}.py" for i in range(16, 29)]

    # Execute each script in order
    for script in scripts_to_run:
        script_path = scripts_directory / script

        # Validate the existence of each script
        if not script_path.exists():
            print(f"Script not found: {script_path}")
            continue

        print(f"Running {script}...")
        try:
            # Execute the script
            result = subprocess.run(
                ["python", str(script_path)],
                capture_output=True,
                text=True,
                check=False,  # Allow the loop to continue even if a script fails
            )

            # Check the result and handle output
            if result.returncode == 0:
                print(f"{script} executed successfully.")
                print(result.stdout)
            else:
                print(f"Error while running {script}:")
                print(result.stderr)

        except Exception as e:
            print(f"Failed to execute {script} due to error: {e}")

    print("All scripts executed.")

if __name__ == "__main__":
    main()
