import os
import subprocess
import logging
from datetime import datetime

# Setup logging
log_file = "MASTER_SCRIPT.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# Dynamic Base Path
base_path = os.getenv("BASE_UPLOAD_PATH", os.path.join("C:", os.sep, "Users", "jaydi", "OneDrive - Comline", "CAPLOCATION", "Deployment", "bulk_report_webapp"))
tasks_path = os.path.join(base_path, "tasks", "TARGETPER")
output_path = os.path.join(base_path, "output", "TARGETPER")
uploads_path = os.path.join(base_path, "uploads", "TARGETPER")

# Ensure required paths exist
required_paths = [tasks_path, output_path, uploads_path]
for path in required_paths:
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
        logging.info(f"Created missing directory: {path}")

# Scripts to run in sequence
scripts = [
    "OUTPUT0.py",  # Step 1: Preprocessing files
    "OUTPUT1.py",  # Step 2: Add columns to TARGET_MODIFIED
    "OUTPUT2.py",  # Step 3: Populate data in OUTPUT2.csv
    "OUTPUT3.py",  # Step 4: Sales After Target
    "OUTPUT4.py",  # Step 5: Summary Generation
    "OUTPUT5.py"   # Step 6: Final Formatting
]

# Function to run each script and log results
def run_script(script_name):
    script_path = os.path.join(tasks_path, script_name)
    if not os.path.exists(script_path):
        logging.error(f"Script not found: {script_path}")
        return False

    try:
        logging.info(f"Running script: {script_name}")
        start_time = datetime.now()
        subprocess.run(["python", script_path], check=True)
        end_time = datetime.now()
        elapsed_time = (end_time - start_time).total_seconds()
        logging.info(f"Successfully executed {script_name} in {elapsed_time:.2f} seconds")
        return True
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running script {script_name}. Details: {e}")
        return False
    except Exception as e:
        logging.error(f"Unexpected error running script {script_name}. Details: {e}")
        return False

# Master script execution
logging.info("Starting MASTER_SCRIPT execution")
for script in scripts:
    if not run_script(script):
        logging.error(f"Aborting execution due to failure in script: {script}")
        break
else:
    logging.info("All scripts executed successfully. MASTER_SCRIPT completed.")

# Summary log message
if os.path.exists(log_file):
    print(f"Execution log available at: {os.path.abspath(log_file)}")
