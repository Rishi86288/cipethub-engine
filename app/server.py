"""
CIPETHUB Engine — Flask API Server
Exposes HTTP endpoints for video generation, download, and status.
"""

import os
import threading
from pathlib import Path

from flask import Flask, jsonify, request, send_file

from generator import generate_full_video

app = Flask(__name__)

# ---------------------------------------------------------------------------
# Shared generation state
# ---------------------------------------------------------------------------

_state = {
    "running": False,
    "last_result": None,
}
_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Endpoints
# ---------------------------------------------------------------------------


@app.get("/health")
def health():
    return jsonify({"status": "ok", "service": "cipethub-engine", "port": 8899})


@app.post("/generate")
def generate():
    with _lock:
        if _state["running"]:
            return jsonify({"error": "Generation already in progress"}), 429
        _state["running"] = True
        _state["last_result"] = None

    body = request.get_json(silent=True) or {}
    topic = body.get("topic", "")
    department = body.get("department", "All")
    video_type = body.get("video_type", "PPT Lecture")

    if not topic:
        with _lock:
            _state["running"] = False
        return jsonify({"error": "Missing required field: topic"}), 400

    valid_depts = {"Plastics", "Mechanical", "Manufacturing", "All"}
    if department not in valid_depts:
        with _lock:
            _state["running"] = False
        return jsonify({"error": f"Invalid department. Must be one of: {valid_depts}"}), 400

    valid_types = {"Simulation", "PPT Lecture"}
    if video_type not in valid_types:
        with _lock:
            _state["running"] = False
        return jsonify({"error": f"Invalid video_type. Must be one of: {valid_types}"}), 400

    try:
        result = generate_full_video(topic, department, video_type)
    except Exception as exc:
        result = {"status": "error", "error": str(exc)}
    finally:
        with _lock:
            _state["running"] = False
            _state["last_result"] = result

    if result.get("status") == "error":
        return jsonify(result), 500

    return jsonify(result), 200


@app.get("/download")
def download():
    """Return the latest generated video file as an MP4 download."""
    videos_dir = Path("/videos")
    if not videos_dir.exists():
        return jsonify({"error": "No videos directory found"}), 404

    mp4_files = sorted(videos_dir.rglob("lecture.mp4"))
    if not mp4_files:
        return jsonify({"error": "No video available for download"}), 404

    latest = mp4_files[-1]
    return send_file(
        str(latest),
        mimetype="video/mp4",
        as_attachment=True,
        download_name="cipethub_lecture.mp4",
    )


@app.get("/status")
def status():
    with _lock:
        return jsonify(
            {
                "running": _state["running"],
                "last_result": _state["last_result"],
            }
        )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    print("=" * 60)
    print("  CIPETHUB Engine running on :8899")
    print("  Automated YouTube Video Generator for CIPET Students")
    print("=" * 60)
    app.run(host="0.0.0.0", port=8899, threaded=True)
