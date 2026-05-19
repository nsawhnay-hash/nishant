import os
import time
import uuid
import threading
from pathlib import Path

import requests
from flask import Flask, render_template, request, jsonify, send_from_directory
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

VIDEOS_DIR = Path("videos")
VIDEOS_DIR.mkdir(exist_ok=True)

HF_TOKEN = os.environ.get("HF_TOKEN", "")
# Free open-source text-to-video model on HF Inference API
MODEL_URL = "https://api-inference.huggingface.co/models/damo-vilab/text-to-video-ms-1.7b"

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
    if not HF_TOKEN:
        return jsonify({"error": "HF_TOKEN not set. Add it to your .env file."}), 500

    job_id = str(uuid.uuid4())
    jobs[job_id] = {"status": "queued", "logs": [], "video_url": None, "error": None}

    thread = threading.Thread(target=_run_generation, args=(job_id, data), daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


def _run_generation(job_id: str, data: dict):
    job = jobs[job_id]
    try:
        job["status"] = "running"

        quality = data.get("quality", "balanced")
        steps_map = {"fast": 15, "balanced": 25, "high": 40}
        steps = steps_map.get(quality, 25)

        res_map = {"256": 256, "512": 512}
        size = res_map.get(data.get("resolution", "256"), 256)

        headers = {
            "Authorization": f"Bearer {HF_TOKEN}",
            "Content-Type": "application/json",
        }
        payload = {
            "inputs": data.get("prompt", ""),
            "parameters": {
                "num_inference_steps": steps,
                "height": size,
                "width": size,
                "num_frames": 16,
            },
        }

        job["logs"].append("Connecting to Hugging Face…")

        # Retry loop — model may need to warm up (returns 503 while loading)
        deadline = time.time() + 600  # 10-min max
        attempt = 0
        resp = None
        while time.time() < deadline:
            attempt += 1
            job["logs"].append(f"Sending request (attempt {attempt})…")
            try:
                resp = requests.post(MODEL_URL, headers=headers, json=payload, timeout=180)
            except requests.Timeout:
                job["logs"].append("Request timed out, retrying…")
                time.sleep(10)
                continue

            if resp.status_code == 200:
                break
            elif resp.status_code == 503:
                try:
                    info = resp.json()
                    est = int(info.get("estimated_time", 20))
                except Exception:
                    est = 20
                wait = min(est, 45)
                job["logs"].append(f"Model is loading on HF servers, waiting {wait}s…")
                time.sleep(wait)
            elif resp.status_code == 429:
                job["logs"].append("Rate limited — waiting 30s…")
                time.sleep(30)
            else:
                try:
                    detail = resp.json()
                except Exception:
                    detail = resp.text[:300]
                raise RuntimeError(f"API error {resp.status_code}: {detail}")

        if resp is None or resp.status_code != 200:
            raise RuntimeError("Model did not respond in time. Please try again.")

        job["logs"].append("Saving video…")
        filename = f"{job_id}.mp4"
        filepath = VIDEOS_DIR / filename
        with open(filepath, "wb") as f:
            f.write(resp.content)

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
    items = [
        {"job_id": jid, "video_url": job["video_url"]}
        for jid, job in jobs.items()
        if job["status"] == "completed" and job["video_url"]
    ]
    return jsonify(items)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
