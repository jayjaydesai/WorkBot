import os
import subprocess
import sys

# Ensure UTF-8 output encoding to avoid UnicodeEncodeError
sys.stdout.reconfigure(encoding='utf-8')

# Define base paths dynamically for Local & Azure compatibility
BASE_DIR = os.getenv("BASE_DIR", os.getcwd())  # Use Azure variable or current directory
UPLOAD_PATH = os.path.join(BASE_DIR, "uploads", "GREPLEN")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")
TASKS_PATH = os.path.abspath(os.path.dirname(__file__))  # Fix path to current directory

# Debugging: Print paths to check correctness
print(f"🔍 DEBUG: TASKS_PATH is set to: {TASKS_PATH}")
print(f"🔍 DEBUG: Looking for scripts in: {TASKS_PATH}")

# Ensure required folders exist
os.makedirs(UPLOAD_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Complete list of scripts in order
scripts = [
    "OUTPUT1.py", "OUTPUT2.py", "OUTPUT3.py"
]

try:
    print("\n\033[96m✅ Starting Replen Backorder Greece Processing...\033[0m\n")  # Cyan color

    for script in scripts:
        script_path = os.path.join(TASKS_PATH, script)

        if os.path.exists(script_path):
            print(f"▶ Running {script}...")  # Show script name before execution

            # Execute script and capture output
            result = subprocess.run(["python", script_path], check=True, capture_output=True, text=True)

            print(f"\033[92m✅ {script} completed successfully!\033[0m")  # Green for success
            if result.stdout:
                print(f"\033[90m{result.stdout.strip()}\033[0m")  # Gray for script output
        else:
            print(f"\033[93m❌ WARNING: {script} not found. Skipping...\033[0m")  # Yellow for warnings

    print("\n\033[92m🎉 Replen Backorder Greece Processing Completed Successfully!\033[0m\n")

except subprocess.CalledProcessError as e:
    print(f"\n\033[91m❌ ERROR in {script}: {e.stderr.strip()}\033[0m\n")  # Red for script failure
except Exception as e:
    print(f"\n\033[91m❌ UNEXPECTED ERROR: {str(e)}\033[0m\n")  # Red for unexpected errors
