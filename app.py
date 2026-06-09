import os
import sys
import uuid
import queue
import threading
import subprocess
from urllib.parse import quote
from flask import Flask, request, Response, send_file, render_template, jsonify, make_response

# Pure Python .env loader
if os.path.exists(".env"):
    with open(".env", "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, val = line.split("=", 1)
                os.environ[key.strip()] = val.strip()

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB max upload

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg"}

# In-memory job store: job_id -> { queue, output_file, status, filename }
jobs: dict = {}


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "image" not in request.files:
        return jsonify({"error": "No file part in request"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Only PNG, JPG, JPEG files are allowed"}), 400

    job_id = str(uuid.uuid4())
    ext = file.filename.rsplit(".", 1)[1].lower()
    saved_name = f"{job_id}.{ext}"
    image_path = os.path.join(UPLOAD_DIR, saved_name)
    file.save(image_path)

    output_xlsx = os.path.splitext(image_path)[0] + ".xlsx"

    # Use user-provided output name, or fall back to original filename stem
    custom_name = request.form.get("output_name", "").strip()
    if custom_name:
        # Sanitize: remove path separators and illegal chars
        safe_chars = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ ()[]")
        custom_name = "".join(c for c in custom_name if c in safe_chars).strip()
    if not custom_name:
        custom_name = os.path.splitext(file.filename)[0]
    download_name = custom_name + ".xlsx"

    q: queue.Queue = queue.Queue()
    jobs[job_id] = {
        "queue": q,
        "output_file": output_xlsx,
        "status": "running",
        "download_name": download_name,
    }

    # Run generate_excel.py in a background thread
    def run_generation():
        try:
            proc = subprocess.Popen(
                [sys.executable, "generate_excel.py", image_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=os.path.dirname(__file__),
            )
            for line in proc.stdout:
                stripped = line.strip()
                if stripped:
                    q.put(("log", stripped))
            proc.wait()
            if proc.returncode == 0 and os.path.exists(output_xlsx):
                jobs[job_id]["status"] = "done"
                q.put(("done", "Excel berhasil digenerate!"))
            else:
                jobs[job_id]["status"] = "error"
                q.put(("error", "Proses generate Excel gagal. Cek terminal untuk detail."))
        except Exception as e:
            jobs[job_id]["status"] = "error"
            q.put(("error", str(e)))
        finally:
            q.put(None)  # sentinel to stop SSE

    t = threading.Thread(target=run_generation, daemon=True)
    t.start()

    return jsonify({"job_id": job_id, "filename": file.filename, "download_name": download_name})


@app.route("/stream/<job_id>")
def stream(job_id: str):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    def event_generator():
        q = jobs[job_id]["queue"]
        while True:
            item = q.get()
            if item is None:
                # End of stream
                yield "event: end\ndata: done\n\n"
                break
            event_type, data = item
            # Escape newlines inside data field
            safe_data = data.replace("\n", " ").replace("\r", "")
            yield f"event: {event_type}\ndata: {safe_data}\n\n"

    return Response(
        event_generator(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
        },
    )


@app.route("/download/<job_id>")
def download(job_id: str):
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    job = jobs[job_id]
    if job["status"] != "done":
        return jsonify({"error": "File not ready yet"}), 400

    output_file = job["output_file"]
    if not os.path.exists(output_file):
        return jsonify({"error": "Output file not found"}), 404

    download_name = job.get("download_name", "output.xlsx")

    # Build response with explicit Content-Disposition to avoid UUID filename bug
    response = make_response(send_file(output_file, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    # RFC 5987 encoding for non-ASCII filenames
    encoded_name = quote(download_name, safe="")
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{download_name}\"; filename*=UTF-8''{encoded_name}"
    )
    return response


if __name__ == "__main__":
    print("=" * 50)
    print("  Image to Excel - Web App")
    print("  Buka browser: http://localhost:5000")
    print("=" * 50)
    app.run(debug=False, port=5000, threaded=True)
