#!/usr/bin/env python3
"""
HTTP endpoint for Microsoft Graph change notifications (webhook).

Graph will NOT call http://localhost — you need a public HTTPS URL (ngrok, Cloudflare Tunnel,
Azure App Service, etc.) pointing at this server, then register:

  POST https://graph.microsoft.com/v1.0/subscriptions
  {
    "changeType": "created",
    "notificationUrl": "https://YOUR-PUBLIC-HOST/graph-notify",
    "resource": "/drives/{RECORDINGS_DRIVE_ID}/items/{RECORDINGS_FOLDER_ID}",
    "expirationDateTime": "...",
    "clientState": "<same as WEBHOOK_CLIENT_STATE in .env>"
  }

Renew subscriptions before they expire (see Graph docs).

GET  /graph-notify?validationToken=...  → plain-text echo (required when creating subscription)
POST /graph-notify                      → JSON notifications; spawns transcribe job in background
GET  /health                            → simple OK for load balancers

Env:
  WEBHOOK_STREAM_TRANSCRIPT=1  → run fetch_sharepoint_stream_transcripts.py (Stream API transcript →
    plain .txt + upload) instead of poll_sharepoint_transcribe.py. Needs CERT_* in .env.
  WEBHOOK_TEAMS_TRANSCRIPT=1   → poll_sharepoint_transcribe.py --teams-transcript (when .vtt exists in folder)
"""

from __future__ import annotations

import logging
import os
import subprocess
import sys
import threading
from pathlib import Path

from dotenv import load_dotenv
from flask import Flask, request

load_dotenv(Path(__file__).resolve().parent / ".env")

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

app = Flask(__name__)


def getenv(name: str) -> str:
    v = os.environ.get(name, "").strip()
    if not v:
        raise RuntimeError(f"Missing env var: {name}")
    return v


def run_poll_for_item(item_id: str) -> None:
    automation_dir = Path(__file__).resolve().parent
    project_root = Path(getenv("PROJECT_ROOT"))
    py = project_root / ".venv" / "bin" / "python"
    if not py.is_file():
        py = Path(sys.executable)

    if os.environ.get("WEBHOOK_STREAM_TRANSCRIPT", "").strip() == "1":
        script = automation_dir / "fetch_sharepoint_stream_transcripts.py"
        cmd = [
            str(py),
            str(script),
            "--file-id",
            item_id,
            "--plain-text",
            "--upload",
            "--no-state",
        ]
    else:
        script = automation_dir / "poll_sharepoint_transcribe.py"
        cmd = [str(py), str(script)]
        if os.environ.get("WEBHOOK_TEAMS_TRANSCRIPT", "").strip() == "1":
            cmd.append("--teams-transcript")
        cmd.extend(["--item-id", item_id, "--keep-local-copy"])

    log.info("Starting background job: %s", " ".join(cmd))
    proc = subprocess.run(
        cmd,
        cwd=str(automation_dir),
        capture_output=True,
        text=True,
        timeout=86400,
    )
    out = (proc.stdout or "") + (proc.stderr or "")
    if proc.returncode != 0:
        log.error("job exit %s output=%s", proc.returncode, out[:3000])
    else:
        log.info("job finished OK for item %s", item_id)
        if out.strip():
            log.info("%s", out.strip()[:2000])


@app.get("/health")
def health():
    return {"status": "ok"}, 200


@app.route("/graph-notify", methods=["GET", "POST"])
def graph_notify():
    # Subscription validation (creation / renewal check)
    if request.method == "GET":
        token = request.args.get("validationToken", "")
        if not token:
            return "", 400
        log.info("Graph validation handshake (GET)")
        return token, 200, {"Content-Type": "text/plain; charset=utf-8"}

    # Notifications
    body = request.get_json(silent=True) or {}
    values = body.get("value") or []
    if not values:
        log.warning("POST /graph-notify with no value[]")
        return "", 202

    expected = os.environ.get("WEBHOOK_CLIENT_STATE", "").strip()
    if not expected:
        log.error("WEBHOOK_CLIENT_STATE not set — refusing notifications (set it to match subscription clientState)")
        return "", 503

    started = 0
    for note in values:
        if note.get("clientState") != expected:
            log.warning("Ignoring notification (clientState mismatch)")
            continue
        change = (note.get("changeType") or "").lower()
        if change not in ("created", "updated"):
            continue
        rd = note.get("resourceData") or {}
        item_id = rd.get("id")
        if not item_id:
            continue
        odata_type = (rd.get("@odata.type") or "").lower()
        if "driveitem" in odata_type or not odata_type:
            pass
        else:
            continue
        log.info("Queue transcribe for driveItem id=%s changeType=%s", item_id, change)
        threading.Thread(
            target=run_poll_for_item,
            args=(item_id,),
            daemon=True,
        ).start()
        started += 1

    if started:
        log.info("Spawned %d background job(s)", started)
    return "", 202


def main() -> int:
    port = int(os.environ.get("WEBHOOK_PORT", "8787"))
    host = os.environ.get("WEBHOOK_HOST", "0.0.0.0")
    log.info("Listening on http://%s:%s  (expose with ngrok/tunnel for Graph)", host, port)
    log.info("Routes: GET/POST /graph-notify  GET /health")
    app.run(host=host, port=port, threaded=True, use_reloader=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
