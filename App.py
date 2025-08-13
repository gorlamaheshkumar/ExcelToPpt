# app.py
import os
import uuid
import subprocess
import threading
import time
import sys
from flask import Flask, request, jsonify, send_from_directory, render_template

# Directories
BASE_DIR = os.path.dirname(__file__)
UPLOAD_DIR = os.path.join(BASE_DIR, "Uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "Outputs")
TEMPLATE_DIR = os.path.join(BASE_DIR, "Templates_Files")  # optional templates
MAIN_SCRIPT = os.path.join(BASE_DIR, "Main.py")  # path to your main.py

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATE_DIR, exist_ok=True)

app = Flask(__name__, static_folder="static", template_folder="Templates")

# In-memory task store (task_id -> info)
tasks = {}

def run_task_in_background(task_id, excel_path, template_path, output_path):
    """
    Spawn main.py with environment variables set. Capture stdout/stderr to a log file.
    The final output will be written to `output_path` (should be a path to Final_Report.pptx).
    """
    task_dir = os.path.dirname(output_path)
    os.makedirs(task_dir, exist_ok=True)

    log_file = os.path.join(task_dir, f"{task_id}.log")
    tasks[task_id]["status"] = "running"
    tasks[task_id]["started_at"] = time.time()
    tasks[task_id]["log_path"] = log_file

    # prepare environment for subprocess
    env = os.environ.copy()
    env["EXCEL_FILE_PATH"] = excel_path
    env["FINAL_OUTPUT_PPTX_PATH"] = output_path

    # Only set TEMPLATE_PPTX_PATH if a template was uploaded; otherwise let main.py
    # fall back to its internal default.
    if template_path:
        env["TEMPLATE_PPTX_PATH"] = template_path

    python_exec = os.environ.get("PYTHON_EXECUTABLE", sys.executable)
    cmd = [python_exec, MAIN_SCRIPT]

    with open(log_file, "wb") as lf:
        try:
            proc = subprocess.Popen(cmd, stdout=lf, stderr=subprocess.STDOUT, env=env)
            ret = proc.wait()
            if ret == 0 and os.path.exists(output_path):
                tasks[task_id]["status"] = "finished"
                tasks[task_id]["finished_at"] = time.time()
                tasks[task_id]["output_path"] = output_path
            else:
                # preserve the log for inspection
                tasks[task_id]["status"] = "failed"
                tasks[task_id]["finished_at"] = time.time()
                tasks[task_id]["error"] = f"Process exited with code {ret}"
        except Exception as e:
            tasks[task_id]["status"] = "failed"
            tasks[task_id]["finished_at"] = time.time()
            tasks[task_id]["error"] = str(e)

@app.route("/")
def index():
    return render_template("Index.html")

@app.route("/upload", methods=["POST"])
def upload_file():
    """
    Accepts:
      - form-data 'file' (required) -> Excel
      - form-data 'template' (optional) -> pptx template
    Creates a task and runs main.py in background.
    Final output path will be: outputs/<task_id>/Final_Report.pptx
    """
    if "file" not in request.files:
        return jsonify({"error": "No file part 'file'"}), 400

    f = request.files["file"]
    if f.filename == "":
        return jsonify({"error": "No selected file"}), 400

    allowed_ext = {".xls", ".xlsx", ".xlsm", ".xlsb"}
    _, ext = os.path.splitext(f.filename.lower())
    if ext not in allowed_ext:
        return jsonify({"error": f"Invalid file type: {ext}. Upload Excel file."}), 400

    task_id = uuid.uuid4().hex
    task_upload_name = f"{task_id}{ext}"
    excel_path = os.path.join(UPLOAD_DIR, task_upload_name)
    f.save(excel_path)

    # Create a per-task output folder and use Final_Report.pptx inside it
    task_output_dir = os.path.join(OUTPUT_DIR, task_id)
    os.makedirs(task_output_dir, exist_ok=True)
    output_path = os.path.join(task_output_dir, "Final_Report.pptx")

    # Optional template upload
    template_file = request.files.get("template")
    template_path = ""
    if template_file and template_file.filename:
        tname = f"{task_id}_template.pptx"
        template_path = os.path.join(TEMPLATE_DIR, tname)
        template_file.save(template_path)
    # If no template uploaded, leave template_path empty so main.py uses its default.

    tasks[task_id] = {
        "status": "queued",
        "uploaded_at": time.time(),
        "excel_path": excel_path,
        "template_path": template_path,
        "output_path": output_path
    }

    # Run in background thread
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
    # send the file from the per-task output directory
    return send_from_directory(os.path.dirname(out), os.path.basename(out), as_attachment=True)

@app.route("/log/<task_id>", methods=["GET"])
def log(task_id):
    t = tasks.get(task_id)
    if not t:
        return jsonify({"error": "task not found"}), 404
    log_path = t.get("log_path")
    if not log_path or not os.path.exists(log_path):
        return jsonify({"error": "log not found"}), 404
    # return tail of the log (last 20000 bytes)
    with open(log_path, "rb") as fh:
        fh.seek(0, os.SEEK_END)
        sz = fh.tell()
        start = max(0, sz - 20000)
        fh.seek(start)
        data = fh.read().decode(errors="ignore")
    return jsonify({"log": data})

if __name__ == "__main__":
    # debug=True for development; change for production
    app.run(debug=True, port=5000, host="0.0.0.0")
