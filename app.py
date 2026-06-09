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

# In-memory job store: job_id -> { queue, output_file, status, download_name }
jobs: dict = {}


# ─── CORS helper ─────────────────────────────────────────────────────────────
def _cors_headers(response):
    """Tambahkan CORS headers agar Vercel frontend bisa akses local API."""
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "*"
    response.headers["Access-Control-Max-Age"] = "3600"
    return response


@app.after_request
def add_cors(response):
    return _cors_headers(response)


@app.route("/v1/", methods=["OPTIONS"])
@app.route("/v1/<path:path>", methods=["OPTIONS"])
def handle_options(path=""):
    """Handle CORS preflight OPTIONS request."""
    resp = make_response("", 204)
    return _cors_headers(resp)


# ─── Helpers ─────────────────────────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def sanitize_name(name: str) -> str:
    safe_chars = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_ ()[]")
    return "".join(c for c in name if c in safe_chars).strip()


# ─── Local Flask UI (localhost only) ─────────────────────────────────────────
@app.route("/")
def index():
    resp = make_response(render_template("index.html"))
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate"
    resp.headers["Pragma"] = "no-cache"
    return resp


# ─── Health check (untuk Vercel frontend detect apakah local server aktif) ──
@app.route("/v1/health")
@app.route("/health")
def health():
    return jsonify({"status": "ok", "service": "image-to-excel"})


# ─── API routes (with /v1 prefix for public tunnel) ──────────────────────────
def _do_upload():
    """Core upload logic, shared by both /upload and /v1/upload."""
    if "image" not in request.files:
        return jsonify({"error": "No file part in request"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Only PNG, JPG, JPEG files are allowed"}), 400

    job_id = str(uuid.uuid4())
    ext = file.filename.rsplit(".", 1)[1].lower()
    image_path = os.path.join(UPLOAD_DIR, f"{job_id}.{ext}")
    file.save(image_path)

    output_xlsx = os.path.splitext(image_path)[0] + ".xlsx"

    custom_name = sanitize_name(request.form.get("output_name", "").strip())
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
                q.put(("error", "Proses generate Excel gagal."))
        except Exception as e:
            jobs[job_id]["status"] = "error"
            q.put(("error", str(e)))
        finally:
            q.put(None)

    threading.Thread(target=run_generation, daemon=True).start()
    return jsonify({"job_id": job_id, "filename": file.filename, "download_name": download_name})


def _do_stream(job_id: str):
    """Core SSE stream logic."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404

    def event_generator():
        q = jobs[job_id]["queue"]
        while True:
            item = q.get()
            if item is None:
                yield "event: end\ndata: done\n\n"
                break
            event_type, data = item
            safe_data = data.replace("\n", " ").replace("\r", "")
            yield f"event: {event_type}\ndata: {safe_data}\n\n"

    return Response(
        event_generator(),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
            "Connection": "keep-alive",
            "Access-Control-Allow-Origin": "*",
        },
    )


def _do_download(job_id: str):
    """Core download logic."""
    if job_id not in jobs:
        return jsonify({"error": "Job not found"}), 404
    job = jobs[job_id]
    if job["status"] != "done":
        return jsonify({"error": "File not ready yet"}), 400
    output_file = job["output_file"]
    if not os.path.exists(output_file):
        return jsonify({"error": "Output file not found"}), 404

    download_name = job.get("download_name", "output.xlsx")
    encoded_name = quote(download_name, safe="")
    response = make_response(send_file(
        output_file,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ))
    response.headers["Content-Disposition"] = (
        f"attachment; filename=\"{download_name}\"; filename*=UTF-8''{encoded_name}"
    )
    return response


# Routes tanpa prefix (untuk localhost testing via Flask UI)
@app.route("/upload", methods=["POST"])
def upload():
    return _do_upload()


@app.route("/stream/<job_id>")
def stream(job_id: str):
    return _do_stream(job_id)


@app.route("/download/<job_id>")
def download(job_id: str):
    return _do_download(job_id)


# Routes dengan /v1 prefix (untuk akses dari Vercel via tunnel)
@app.route("/v1/upload", methods=["POST"])
def v1_upload():
    return _do_upload()


@app.route("/v1/stream/<job_id>")
def v1_stream(job_id: str):
    return _do_stream(job_id)


@app.route("/v1/download/<job_id>")
def v1_download(job_id: str):
    return _do_download(job_id)


if __name__ == "__main__":
    print("=" * 50)
    print("  Image to Excel - Web App")
    print("  Local UI : http://localhost:5000")
    print("  API v1   : http://localhost:5000/v1")
    print("  Tunnel   : https://r3huyik.abc-tunnel.us/v1")
    print("=" * 50)
    app.run(debug=False, port=5000, threaded=True)
