"""
AutoDoc - Flask Web Server
Exposes REST API for the front-end UI and handles file uploads.
"""

import os
import json
import csv
import datetime
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from werkzeug.utils import secure_filename
from autodoc_engine import run_pipeline, LOG_FILE

app = Flask(__name__, static_folder="frontend", static_url_path="")
CORS(app)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
ALLOWED_EXTENSIONS = {"xlsx", "xls"}


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


# ── Static frontend ────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("frontend", "index.html")


# ── Upload & Generate ─────────────────────────────────────────────────────────
@app.route("/api/generate", methods=["POST"])
def generate():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    project = request.form.get("project", "Unnamed Project")
    engineer = request.form.get("engineer", "AutoDoc User")

    if file.filename == "" or not allowed_file(file.filename):
        return jsonify({"error": "Invalid file. Please upload an .xlsx file."}), 400

    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    result = run_pipeline(filepath, project, engineer)
    return jsonify(result)


# ── Download Report ────────────────────────────────────────────────────────────
@app.route("/api/download/<path:filename>")
def download(filename):
    safe = secure_filename(os.path.basename(filename))
    return send_file(os.path.join("generated_reports", safe),
                     as_attachment=True, download_name=safe)


# ── Log / Dashboard Data ───────────────────────────────────────────────────────
@app.route("/api/logs")
def logs():
    if not os.path.isfile(LOG_FILE):
        return jsonify([])
    rows = []
    with open(LOG_FILE, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return jsonify(rows)


@app.route("/api/stats")
def stats():
    if not os.path.isfile(LOG_FILE):
        return jsonify({
            "total": 0, "success": 0, "error": 0,
            "total_components": 0, "by_day": [], "by_project": []
        })

    rows = []
    with open(LOG_FILE, newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    total     = len(rows)
    success   = sum(1 for r in rows if r.get("status") == "Success")
    error     = total - success
    total_comp = sum(int(r.get("component_count", 0) or 0) for r in rows)

    # Group by date
    by_day = {}
    for r in rows:
        day = r.get("generated_at", "")[:10]
        by_day[day] = by_day.get(day, 0) + 1

    # Group by project
    by_project = {}
    for r in rows:
        p = r.get("project", "Unknown")
        by_project[p] = by_project.get(p, 0) + 1

    return jsonify({
        "total":            total,
        "success":          success,
        "error":            error,
        "total_components": total_comp,
        "by_day":           [{"date": k, "count": v} for k, v in sorted(by_day.items())],
        "by_project":       [{"project": k, "count": v} for k, v in by_project.items()],
    })


# ── Sample Excel download ──────────────────────────────────────────────────────
@app.route("/api/sample")
def sample():
    """Returns a sample Excel template for users to fill in."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Components"
    headers = ["Component ID", "Name", "Type", "Voltage Rating (V)",
               "Current Rating (A)", "Material", "Status", "Engineer", "Notes"]
    ws.append(headers)
    sample_rows = [
        ["C-001", "Main Breaker", "Circuit Breaker", 480, 100, "Steel", "Approved", "J. Smith", "UL Listed"],
        ["C-002", "Bus Bar L1",   "Bus Bar",         480,  200, "Copper","Under Review","A. Patel","Check torque spec"],
        ["C-003", "Control Relay","Relay",           24,   5,   "Plastic","Approved","J. Smith","DIN rail mount"],
        ["C-004", "Earth Ground", "Grounding",       0,    0,   "Copper","Pending","B. Lee","Awaiting drawing"],
        ["C-005", "Terminal Block","Terminal",        600,  30,  "Nylon", "Approved","A. Patel","Phoenix Contact"],
    ]
    for row in sample_rows:
        ws.append(row)
    path = "uploads/sample_components.xlsx"
    wb.save(path)
    return send_file(path, as_attachment=True, download_name="sample_components.xlsx")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
