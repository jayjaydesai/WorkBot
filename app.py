from flask import Flask, render_template, request, send_file, flash
from dotenv import load_dotenv
from pathlib import Path
import os
import subprocess  # Import subprocess to run external scripts
from tasks.ALLOUTPUTS import run_all_outputs_for_all_aisles  # Import your workflow script

# Define the Flask app
app = Flask(__name__)

# Define the base directory (root of your app)
BASE_DIR = Path(__file__).resolve().parent

app.secret_key = "your_secret_key"  # Replace with a secure random key

# Load environment variables from .env file
load_dotenv()

# Define folder paths from environment variables
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "uploads")
UPLOAD_FOLDER_REPLEN = BASE_DIR / "uploads" / "REPLEN"
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "output")
OUTPUT_FOLDER_REPLEN = BASE_DIR / "output" / "REPLEN"
TASKS_FOLDER = Path(os.getenv("TASKS_FOLDER", BASE_DIR / "tasks" / "REPLEN"))

# Create necessary folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
UPLOAD_FOLDER_REPLEN.mkdir(parents=True, exist_ok=True)
OUTPUT_FOLDER_REPLEN.mkdir(parents=True, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    """Render the home page and handle file uploads."""
    if request.method == "POST":
        file = request.files.get("file")
        aisle_id = request.form.get("aisle")  # Get aisle ID from the form input

        if file:
            # Save the uploaded BULK.xlsx file to the 'uploads' folder
            bulk_file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(bulk_file_path)

            try:
                # If no aisle ID is provided, process all aisles
                aisle_ids = [aisle_id] if aisle_id else None

                # Run the workflow
                results = run_all_outputs_for_all_aisles(bulk_file_path, aisle_ids)

                # Generate HTML response with download links
                response_html = """
                <h2>Processing Complete</h2>
                <ul>
                """
                for aisle, path in results.items():
                    response_html += f"<li>Aisle {aisle}: <a href='/download/{os.path.basename(path)}'>{os.path.basename(path)}</a></li>"
                response_html += "</ul>"

                # Add "Back to Home" button
                response_html += """
                <br><a href="/" style="text-decoration:none; padding:10px; background-color:blue; color:white; border-radius:5px;">Back to Home</a>
                """
                return response_html

            except Exception as e:
                # Return an error message in case of exceptions
                return f"""
                <h2>An error occurred:</h2>
                <p>{str(e)}</p>
                <br><a href="/" style="text-decoration:none; padding:10px; background-color:red; color:white; border-radius:5px;">Back to Home</a>
                """

    # Render the home page
    return render_template("index.html")

@app.route("/run-replen", methods=["POST"])
def run_replen():
    """Handle REPLEN Stock task."""
    try:
        # Get the uploaded files
        uploaded_files = request.files.getlist("replen_files[]")
        
        # Ensure files were uploaded
        if not uploaded_files:
            flash("No files were uploaded for REPLEN task.", "error")
            return render_template("index.html")

        # Save files to the REPLEN upload folder
        for file in uploaded_files:
            if file.filename.endswith(".xlsx"):
                file_path = UPLOAD_FOLDER_REPLEN / file.filename
                file.save(file_path)
            else:
                flash(f"Invalid file type: {file.filename}. Only .xlsx files are allowed.", "error")
                return render_template("index.html")

        # Run the REPLEN master script
        master_script = TASKS_FOLDER / "REPLEN_ALLOUTPUTS.py"
        if not master_script.exists():
            flash("REPLEN master script not found.", "error")
            return render_template("index.html")

        subprocess.run(
            ["python", str(master_script)],
            check=True,
            text=True
        )

        # Check for the final output file (OUTPUT28.xlsx)
        output_file = OUTPUT_FOLDER_REPLEN / "OUTPUT28.xlsx"
        if output_file.exists():
            download_link = f"/download/replen/{output_file.name}"
            flash("REPLEN task completed successfully.", "success")
            return render_template(
                "index.html",
                download_link=download_link
            )
        else:
            flash("REPLEN task completed, but OUTPUT28.xlsx was not generated.", "error")
            return render_template("index.html")

    except subprocess.CalledProcessError as e:
        flash(f"Error while running REPLEN master script: {e.stderr}", "error")
        return render_template("index.html")
    except Exception as e:
        flash(f"Unexpected error: {str(e)}", "error")
        return render_template("index.html")

@app.route("/download/<filename>")
def download_file(filename):
    """Serve processed files for download."""
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        return send_file(file_path, as_attachment=True)
    except FileNotFoundError:
        # Handle missing file scenario
        return f"""
        <h2>File not found:</h2>
        <p>{filename}</p>
        <br><a href="/" style="text-decoration:none; padding:10px; background-color:red; color:white; border-radius:5px;">Back to Home</a>
        """
@app.route("/download/replen/<filename>")
def download_replen_file(filename):
    """Serve REPLEN output files for download."""
    try:
        file_path = OUTPUT_FOLDER_REPLEN / filename
        if file_path.exists():
            return send_file(file_path, as_attachment=True)
        else:
            flash(f"File not found: {filename}", "error")
            return render_template("index.html")
    except Exception as e:
        flash(f"Unexpected error: {str(e)}", "error")
        return render_template("index.html")

@app.route("/run-checkempty", methods=["POST"])
def run_checkempty():
    try:
        # Log environment and file upload
        print("Starting CHECKEMPTY Task...")
        print(f"Environment Variables: {os.environ}")
        
        # Check if a file is uploaded
        file = request.files.get("file")
        if not file:
            print("No file uploaded.")
            return """
            <h2>Error: No file uploaded.</h2>
            <br><a href="/" style="text-decoration:none; padding:10px; background-color:red; color:white; border-radius:5px;">Back to Home</a>
            """

        # Save the uploaded BULK.xlsx file to the 'uploads' folder
        upload_folder = os.path.abspath(UPLOAD_FOLDER)
        os.makedirs(upload_folder, exist_ok=True)  # Ensure the folder exists
        file_path = os.path.join(upload_folder, "BULK.xlsx")
        file.save(file_path)
        print(f"Uploaded BULK.xlsx saved to: {file_path}")

        # Run the scripts
        scripts = ["OUTPUT11.py", "OUTPUT12.py", "OUTPUT13.py"]
        base_path = os.path.abspath(os.path.join("tasks"))  # Absolute path to the tasks folder

        for script in scripts:
            script_path = os.path.join(base_path, script)
            print(f"Running script: {script_path}")
            result = subprocess.run(
                ["python", script_path],
                capture_output=True,
                text=True,
                check=True
            )
            print(f"Output from {script}: {result.stdout}")

        # Return a success message with the download link
        print("CHECKEMPTY Task completed successfully.")
        return f"""
        <h2>Request Processed Successfully</h2>
        <ul>
            <li><a href="/download/EMPTYLOCATION.xlsx" target="_blank">Download EMPTYLOCATION.xlsx</a></li>
        </ul>
        <br><a href="/" style="text-decoration:none; padding:10px; background-color:blue; color:white; border-radius:5px;">Back to Home</a>
        """
    except subprocess.CalledProcessError as e:
        print(f"Subprocess error while running scripts: {e.stderr}")
        return f"""
        <h2>Error during script execution:</h2>
        <pre>{e.stderr}</pre>
        <br><a href="/" style="text-decoration:none; padding:10px; background-color:red; color:white; border-radius:5px;">Back to Home</a>
        """
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return f"""
        <h2>An unexpected error occurred:</h2>
        <pre>{str(e)}</pre>
        <br><a href="/" style="text-decoration:none; padding:10px; background-color:red; color:white; border-radius:5px;">Back to Home</a>
        """

if __name__ == "__main__":
    app.run(debug=True)

