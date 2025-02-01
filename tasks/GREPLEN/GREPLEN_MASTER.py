import os
import subprocess

# Define base paths dynamically
BASE_DIR = os.getenv("BASE_DIR", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
UPLOAD_PATH = os.path.join(BASE_DIR, "uploads", "GREPLEN")
OUTPUT_PATH = os.path.join(BASE_DIR, "output", "GREPLEN")
TASKS_PATH = os.path.join(BASE_DIR, "tasks", "GREPLEN")

# Ensure required folders exist
os.makedirs(UPLOAD_PATH, exist_ok=True)
os.makedirs(OUTPUT_PATH, exist_ok=True)

# Complete list of scripts in order
scripts = [
    "OUTPUT1.py", "OUTPUT2.py", "OUTPUT3.py", "OUTPUT4.py", "OUTPUT5.py",
    "OUTPUT6.py", "OUTPUT7.py", "OUTPUT8.py", "OUTPUT9.py", "OUTPUT10.py",
    "OUTPUT11.py", "OUTPUT12.py", "OUTPUT13.py", "OUTPUT14.py", "OUTPUT15.py",
    "OUTPUT16.py", "OUTPUT17.py", "OUTPUT18.py", "OUTPUT19.py", "OUTPUT20.py",
    "OUTPUT21.py", "OUTPUT22.py", "OUTPUT23.py", "OUTPUT24.py", "OUTPUT25.py",
    "OUTPUT26.py", "OUTPUT27.py"
]

try:
    print("\n✅ Starting Replen Backorder Greece Processing...\n")

    for script in scripts:
        script_path = os.path.join(TASKS_PATH, script)

        if os.path.exists(script_path):
            print(f"▶ Running {script}...")
            result = subprocess.run(["python", script_path], check=True, text=True, capture_output=True)
            print(result.stdout)  # Print script output
        else:
            print(f"❌ WARNING: {script} not found. Skipping...")

    print("\n🎉 Replen Backorder Greece Processing Completed Successfully!\n")

except subprocess.CalledProcessError as e:
    print(f"\n❌ ERROR: Script execution failed - {e.stderr}\n")
except Exception as e:
    print(f"\n❌ UNEXPECTED ERROR: {str(e)}\n")
