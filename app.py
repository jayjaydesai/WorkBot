from flask import Flask, render_template, request, send_file
from dotenv import load_dotenv
import os
from tasks.ALLOUTPUTS import run_all_outputs_for_all_aisles  # Import your workflow script

# Define the Flask app
app = Flask(__name__)

# Load environment variables from .env file
load_dotenv()

# Define folder paths from environment variables
UPLOAD_FOLDER = os.getenv("UPLOAD_FOLDER", "uploads")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "output")

# Create necessary folders if they don't exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

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

if __name__ == "__main__":
    # Use host='0.0.0.0' for deployment and disable debug mode in production
    debug_mode = os.getenv("FLASK_DEBUG", "True").lower() == "true"
    app.run(debug=debug_mode, host="0.0.0.0")
