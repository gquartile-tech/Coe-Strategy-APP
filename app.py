"""
CoE Account Strategy Analysis Tool — Flask backend
Mirrors the structure of the CoE Account Mastery Analysis app.
"""

from __future__ import annotations

import gc
import os
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request
from werkzeug.utils import secure_filename

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR      = Path(__file__).parent.resolve()
UPLOAD_DIR    = BASE_DIR / "uploads"
OUTPUT_DIR    = BASE_DIR / "outputs"
TEMPLATE_FILE = BASE_DIR / "CoE_Account_Strategy_Analysis_Templates_V2.xlsm"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

sys.path.insert(0, str(BASE_DIR))

from writer_strategy import write_strategy

MIN_OUTPUT_BYTES = 5_000

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


def run_full_analysis(input_path: str) -> dict:
    if not TEMPLATE_FILE.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_FILE}")

    result_path   = write_strategy(str(input_path), str(TEMPLATE_FILE), str(OUTPUT_DIR))
    download_path = Path(result_path)

    size = download_path.stat().st_size if download_path.exists() else 0
    if not download_path.exists() or size < MIN_OUTPUT_BYTES:
        raise RuntimeError(f"Output file missing or too small ({size} bytes).")

    # Parse account label and date range from filename
    # Pattern: "{label} — Strategy Analysis {date_range}.xlsm"
    fname   = download_path.stem
    parts   = fname.split(" \u2014 Strategy Analysis ")
    account = parts[0].strip() if parts else fname
    window  = parts[1].strip() if len(parts) > 1 else ""

    return {
        "download_filename": download_path.name,
        "account":           account,
        "window":            window,
    }


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/analyze", methods=["POST"])
def analyze():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400
    uploaded = request.files["file"]
    if not uploaded.filename:
        return jsonify({"error": "No file selected."}), 400
    _, ext = os.path.splitext(uploaded.filename.lower())
    if ext not in {".xlsx", ".xlsm"}:
        return jsonify({"error": "Only .xlsx or .xlsm files accepted."}), 400

    safe_name  = secure_filename(uploaded.filename)
    input_path = UPLOAD_DIR / safe_name
    uploaded.save(str(input_path))

    try:
        info = run_full_analysis(str(input_path))
    except FileNotFoundError as e:
        return jsonify({"error": str(e)}), 500
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": f"Analysis failed: {e}"}), 500
    finally:
        gc.collect()

    info["download_url"] = f"/download/{info['download_filename']}"
    return jsonify(info)


@app.route("/download/<path:filename>")
def download(filename):
    from urllib.parse import unquote
    filename = unquote(filename)
    p = OUTPUT_DIR / filename

    if not p.exists():
        xlsm_files = sorted(OUTPUT_DIR.glob("*.xlsm"), key=lambda f: f.stat().st_mtime, reverse=True)
        if xlsm_files:
            p        = xlsm_files[0]
            filename = p.name
        else:
            return f"No output files found in {OUTPUT_DIR}", 404

    data = p.read_bytes()
    return Response(
        data,
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length":      str(len(data)),
        },
    )


@app.route("/healthcheck")
def healthcheck():
    template_ok = TEMPLATE_FILE.exists()
    return jsonify({
        "status":             "ok" if template_ok else "degraded",
        "agent":              "account_strategy",
        "template_reachable": template_ok,
    }), 200 if template_ok else 503


@app.route("/favicon.ico")
def favicon():
    return "", 204


if __name__ == "__main__":
    print("\n  CoE Account Strategy Analysis Tool")
    print("  ─────────────────────────────────────────────────")
    print(f"  Template : {TEMPLATE_FILE}")
    print(f"  Template exists: {TEMPLATE_FILE.exists()}")
    print(f"  Outputs  : {OUTPUT_DIR}")
    print("  Open → http://127.0.0.1:8504\n")
    app.run(host="127.0.0.1", port=8504, debug=True)
