import os
import uuid
import subprocess
import threading
import time
import sys
from flask import Flask, request, jsonify, send_from_directory, render_template

# --- DYNAMIC PATH SETUP ---
# This section correctly calculates all paths based on the script's location.

# Get the directory of the current script (e.g., .../ExcelToPpt/Linux)
script_dir = os.path.dirname(os.path.abspath(__file__))
# Get the parent directory, which is the project's root (e.g., .../ExcelToPpt)
BASE_DIR = os.path.dirname(script_dir)

# Define all other paths relative to the project's root directory
UPLOAD_DIR = os.path.join(BASE_DIR, "Uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "Outputs")
TEMPLATE_INPUT_DIR = os.path.join(BASE_DIR, "Templates_Files")
MAIN_SCRIPT = os.path.join(script_dir, "Main_Linux.py") # Main_Linux.py is in the same folder as this script

# Define the paths for Flask's static and template folders
TEMPLATE_FOLDER = os.path.join(BASE_DIR, "Templates")
STATIC_FOLDER = os.path.join(BASE_DIR, "static")

# --- END OF PATH SETUP ---

# Create directories if they don't exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATE_INPUT_DIR, exist_ok=True)

# Initialize Flask, telling it where to find the templates folder
app = Flask(__name__, template_folder=TEMPLATE_FOLDER, static_folder=STATIC_FOLDER)

# In-memory task store (task_id -> info)
tasks = {}

def run_task_in_background(task_id, excel_path, template_path, output_path):
    """
    Spawns Main_Linux.py with environment variables set and captures logs.
    """
    log_file = os.path.join(os.path.dirname(output_path), f"{task_id}.log")
    tasks[task_id]["status"] = "running"
    tasks[task_id]["started_at"] = time.time()
    tasks[task_id]["log_path"] = log_file

    env = os.environ.copy()
    env["EXCEL_FILE_PATH"] = excel_path
    env["FINAL_OUTPUT_PPTX_PATH"] = output_path
    
    if template_path:
        env["TEMPLATE_PPTX_PATH"] = template_path

    python_exec = sys.executable
    cmd = [python_exec, MAIN_SCRIPT]

    with open(log_file, "wb") as lf:
        try:
            proc = subprocess.Popen(cmd, stdout=lf, stderr=subprocess.STDOUT, env=env)
            ret = proc.wait()
            
            if ret == 0 and os.path.exists(output_path):
                tasks[task_id]["status"] = "finished"
            else:
                tasks[task_id]["status"] = "failed"
                tasks[task_id]["error"] = f"Process exited with code {ret}. Check log for details."
        except Exception as e:
            tasks[task_id]["status"] = "failed"
            tasks[task_id]["error"] = str(e)
        finally:
            tasks[task_id]["finished_at"] = time.time()

@app.route("/")
def index():
    return render_template("Index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file part 'file'"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "No selected file"}), 400

    task_id = uuid.uuid4().hex
    _, ext = os.path.splitext(f.filename.lower())
    task_upload_name = f"{task_id}{ext}"
    excel_path = os.path.join(UPLOAD_DIR, task_upload_name)
    f.save(excel_path)

    task_output_dir = os.path.join(OUTPUT_DIR, task_id)
    os.makedirs(task_output_dir, exist_ok=True)
    output_path = os.path.join(task_output_dir, "Final_Report.pptx")

    template_file = request.files.get("template")
    template_path = ""
    if template_file and template_file.filename:
        tname = f"{task_id}_template.pptx"
        template_path = os.path.join(TEMPLATE_INPUT_DIR, tname)
        template_file.save(template_path)

    tasks[task_id] = {
        "status": "queued",
        "uploaded_at": time.time(),
        "excel_path": excel_path,
        "template_path": template_path,
        "output_path": output_path
    }

    thread = threading.Thread(target=run_task_in_background, args=(task_id, excel_path, template_path, output_path))
    thread.daemon = True
    thread.start()

    return jsonify({"task_id": task_id}), 202

@app.route("/status/<task_id>", methods=["GET"])
def status(task_id):
    t = tasks.get(task_id)
    if not t:
        return jsonify({"error": "task not found"}), 404
    
    response = {
        "task_id": task_id,
        "status": t.get("status"),
        "uploaded_at": t.get("uploaded_at"),
        "started_at": t.get("started_at"),
        "finished_at": t.get("finished_at"),
        "error": t.get("error"),
        "has_output": os.path.exists(t.get("output_path", "")) if t.get("output_path") else False,
        "log_url": f"/log/{task_id}"
    }
    return jsonify(response)

@app.route("/download/<task_id>", methods=["GET"])
def download(task_id):
    t = tasks.get(task_id)
    if not t:
        return jsonify({"error": "task not found"}), 404
    out = t.get("output_path")
    if not out or not os.path.exists(out):
        return jsonify({"error": "output not ready"}), 404
    return send_from_directory(os.path.dirname(out), os.path.basename(out), as_attachment=True)

@app.route("/log/<task_id>", methods=["GET"])
def log(task_id):
    t = tasks.get(task_id)
    if not t:
        return jsonify({"error": "task not found"}), 404
    log_path = t.get("log_path")
    if not log_path or not os.path.exists(log_path):
        return jsonify({"error": "log not found"}), 404

    with open(log_path, "rb") as fh:
        fh.seek(0, os.SEEK_END)
        sz = fh.tell()
        start = max(0, sz - 20000)
        fh.seek(start)
        data = fh.read().decode(errors="ignore")
    return jsonify({"log": data})

if __name__ == "__main__":
    app.run(debug=True, port=5000, host="0.0.0.0")