import os
import uuid
import threading
import requests
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_from_directory
import fal_client

app = Flask(__name__)

VIDEOS_DIR = Path("videos")
VIDEOS_DIR.mkdir(exist_ok=True)

# In-memory job store (sufficient for personal use)
jobs: dict = {}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    data = request.json or {}
    prompt = data.get("prompt", "").strip()
    if not prompt:
        return jsonify({"error": "Prompt is required"}), 400

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "queued", "logs": [], "video_url": None, "error": None}

    thread = threading.Thread(target=_run_generation, args=(job_id, data), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


def _run_generation(job_id: str, data: dict):
    job = jobs[job_id]
    try:
        job["status"] = "running"
        job["logs"].append("Submitting to Kling AI...")

        model = "fal-ai/kling-video/v1.6/pro/text-to-video"
        args = {
            "prompt": data.get("prompt", ""),
            "negative_prompt": data.get("negative_prompt", ""),
            "duration": data.get("duration", "5"),
            "aspect_ratio": data.get("aspect_ratio", "16:9"),
            "cfg_scale": float(data.get("cfg_scale", 0.5)),
        }

        def on_update(update):
            if isinstance(update, fal_client.InProgress):
                for log in update.logs:
                    msg = log.get("message", "").strip()
                    if msg:
                        job["logs"].append(msg)

        result = fal_client.subscribe(
            model,
            arguments=args,
            with_logs=True,
            on_queue_update=on_update,
        )

        video_url = result["video"]["url"]
        job["logs"].append("Downloading video...")

        filename = f"{job_id}.mp4"
        filepath = VIDEOS_DIR / filename
        resp = requests.get(video_url, stream=True, timeout=120)
        resp.raise_for_status()
        with open(filepath, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)

        job["status"] = "completed"
        job["video_url"] = f"/videos/{filename}"
        job["logs"].append("Done! Your video is ready.")

    except Exception as exc:
        job["status"] = "failed"
        job["error"] = str(exc)
        job["logs"].append(f"Error: {exc}")


@app.route("/status/<job_id>")
def status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/videos/<filename>")
def serve_video(filename: str):
    return send_from_directory(VIDEOS_DIR, filename)


@app.route("/history")
def history():
    items = []
    for jid, job in jobs.items():
        if job["status"] == "completed" and job["video_url"]:
            items.append({"job_id": jid, "video_url": job["video_url"]})
    return jsonify(items)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
