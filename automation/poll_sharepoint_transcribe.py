#!/usr/bin/env python3
"""
Poll a SharePoint folder (e.g. Teams Recordings):

  Default: new media files (.mp4, …) → download → local Whisper (../transcribe.py) → upload .txt

  --teams-transcript: new Teams transcript files (.vtt / .txt from the meeting) → download →
  convert VTT to plain text → upload .txt (no local transcription).

Teams must be saving transcripts into that folder (organizer / policy). Often a .vtt appears
next to the recording after the meeting.

Transcripts are often stored in **subfolders** under Recordings (Teams date/title folders). By default
this script only lists **immediate** children of RECORDINGS_FOLDER_ID. Use --recursive or
RECORDINGS_RECURSIVE=1 to discover .vtt/.txt in nested folders.

Azure AD app registration (application permissions, admin consent):
  - Files.Read.All, Sites.ReadWrite.All, optional Mail.Send

https://learn.microsoft.com/en-us/graph/auth-v2-service
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv(Path(__file__).resolve().parent / ".env")

GRAPH = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]
RECORDING_SUFFIXES = (".mp4", ".m4a", ".mp3", ".wav", ".aac", ".mov")
# Teams / SharePoint transcript artifacts (--teams-transcript). Override: TEAMS_TRANSCRIPT_EXTENSIONS=.vtt,.txt
DEFAULT_TRANSCRIPT_SUFFIXES = (".vtt", ".txt")

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
log = logging.getLogger(__name__)

STATE_PATH = Path(__file__).resolve().parent / "processed_recordings.json"
STATE_PATH_TEAMS_TRANSCRIPTS = (
    Path(__file__).resolve().parent / "processed_teams_transcripts.json"
)


def transcript_suffixes() -> tuple[str, ...]:
    raw = os.environ.get("TEAMS_TRANSCRIPT_EXTENSIONS", "").strip().lower()
    if not raw:
        return DEFAULT_TRANSCRIPT_SUFFIXES
    parts = tuple(p.strip() for p in raw.split(",") if p.strip())
    if not parts:
        return DEFAULT_TRANSCRIPT_SUFFIXES
    out = []
    for p in parts:
        if not p.startswith("."):
            p = "." + p
        out.append(p)
    return tuple(out)


def getenv(name: str, default: str | None = None) -> str:
    v = os.environ.get(name, default)
    if v is None or v == "":
        raise SystemExit(f"Missing required env var: {name}")
    return v


def transcripts_drive_id() -> str:
    v = os.environ.get("TRANSCRIPTS_DRIVE_ID", "").strip()
    if v:
        return v
    log.info("TRANSCRIPTS_DRIVE_ID unset; using RECORDINGS_DRIVE_ID for transcript upload")
    return getenv("RECORDINGS_DRIVE_ID")


def transcripts_folder_id() -> str:
    v = os.environ.get("TRANSCRIPTS_FOLDER_ID", "").strip()
    if v:
        return v
    log.info("TRANSCRIPTS_FOLDER_ID unset; using RECORDINGS_FOLDER_ID for transcript upload")
    return getenv("RECORDINGS_FOLDER_ID")


def graph_token() -> str:
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    result = app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description") or str(result))
    return result["access_token"]


def load_processed(state_path: Path) -> set[str]:
    if not state_path.is_file():
        return set()
    data = json.loads(state_path.read_text(encoding="utf-8"))
    return set(data.get("drive_item_ids", []))


def save_processed(ids: set[str], state_path: Path) -> None:
    state_path.write_text(
        json.dumps({"drive_item_ids": sorted(ids)}, indent=2),
        encoding="utf-8",
    )


def _drive_children_pages(token: str, drive: str, folder_id: str):
    url = f"{GRAPH}/drives/{drive}/items/{folder_id}/children"
    while url:
        r = httpx.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=120.0,
        )
        r.raise_for_status()
        body = r.json()
        for it in body.get("value", []):
            yield it
        url = body.get("@odata.nextLink")


def _recursive_max_depth() -> int:
    raw = os.environ.get("RECORDINGS_RECURSIVE_MAX_DEPTH", "16").strip()
    try:
        d = int(raw)
    except ValueError:
        return 16
    return max(1, min(d, 50))


def list_folder_files(
    token: str, suffixes: tuple[str, ...], *, recursive: bool = False
) -> list[dict]:
    drive = getenv("RECORDINGS_DRIVE_ID")
    root = getenv("RECORDINGS_FOLDER_ID")
    items: list[dict] = []

    if not recursive:
        for it in _drive_children_pages(token, drive, root):
            if "folder" in it:
                continue
            name = (it.get("name") or "").lower()
            if any(name.endswith(s) for s in suffixes):
                items.append(it)
        return items

    max_depth = _recursive_max_depth()

    def walk(folder_id: str, depth: int) -> None:
        if depth > max_depth:
            return
        for it in _drive_children_pages(token, drive, folder_id):
            if "folder" in it:
                walk(it["id"], depth + 1)
                continue
            name = (it.get("name") or "").lower()
            if any(name.endswith(s) for s in suffixes):
                items.append(it)

    walk(root, 0)
    return items


def list_folder_all_files(token: str, *, recursive: bool = False) -> list[dict]:
    """Files under RECORDINGS folder; use recursive=True to include nested folders."""
    drive = getenv("RECORDINGS_DRIVE_ID")
    root = getenv("RECORDINGS_FOLDER_ID")
    items: list[dict] = []

    if not recursive:
        for it in _drive_children_pages(token, drive, root):
            if "folder" in it:
                continue
            items.append(it)
        return items

    max_depth = _recursive_max_depth()

    def walk(folder_id: str, depth: int) -> None:
        if depth > max_depth:
            return
        for it in _drive_children_pages(token, drive, folder_id):
            if "folder" in it:
                walk(it["id"], depth + 1)
            else:
                items.append(it)

    walk(root, 0)
    return items


def vtt_to_plain_text(vtt: str) -> str:
    """Strip WEBVTT cues/timestamps; keep readable lines (good enough for meeting notes)."""
    out_lines: list[str] = []
    for raw in vtt.splitlines():
        line = raw.strip()
        if not line:
            continue
        low = line.lower()
        if low.startswith("webvtt") or low.startswith("note ") or low.startswith("kind:"):
            continue
        if low.startswith("language:"):
            continue
        if "-->" in line:
            continue
        if line.isdigit():
            continue
        if line.startswith("STYLE") or line.startswith("REGION"):
            continue
        out_lines.append(line)
    return "\n".join(out_lines).strip() + "\n"


def get_drive_item(token: str, drive_id: str, item_id: str) -> dict:
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}"
    r = httpx.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        timeout=60.0,
    )
    r.raise_for_status()
    return r.json()


def download_item(token: str, drive_id: str, item_id: str, dest: Path) -> None:
    url = f"{GRAPH}/drives/{drive_id}/items/{item_id}/content"
    with httpx.stream(
        "GET",
        url,
        headers={"Authorization": f"Bearer {token}"},
        follow_redirects=True,
        timeout=600.0,
    ) as r:
        r.raise_for_status()
        with dest.open("wb") as f:
            for chunk in r.iter_bytes(1024 * 1024):
                f.write(chunk)


def upload_transcript(
    token: str, drive_id: str, parent_id: str, filename: str, body: bytes
) -> None:
    enc = quote(filename, safe="!$&'()*+,;=:@-._~ ")
    url = f"{GRAPH}/drives/{drive_id}/items/{parent_id}:/{enc}:/content"
    r = httpx.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "text/plain; charset=utf-8",
        },
        content=body,
        timeout=120.0,
    )
    r.raise_for_status()


def s3_put_if_configured(local_path: Path, key: str) -> None:
    bucket = os.environ.get("S3_BUCKET", "").strip()
    if not bucket:
        log.info("S3_BUCKET not set; skipping S3")
        return
    import boto3

    region = os.environ.get("AWS_REGION", "us-east-1")
    boto3.client("s3", region_name=region).upload_file(
        str(local_path), bucket, key, ExtraArgs={"ContentType": "text/plain; charset=utf-8"}
    )
    log.info("Uploaded to s3://%s/%s", bucket, key)


def send_mail_graph(token: str, subject: str, body_text: str) -> None:
    if os.environ.get("MAIL_ON_SHAREPOINT_FAILURE", "0") != "1":
        return
    sender = os.environ.get("MAIL_SENDER_USER_ID", "").strip()
    to_addr = os.environ.get("MAIL_TO", "").strip()
    if not sender or not to_addr:
        log.warning("Mail fallback enabled but MAIL_SENDER_USER_ID or MAIL_TO missing")
        return
    url = f"{GRAPH}/users/{sender}/sendMail"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": to_addr}}],
        },
        "saveToSentItems": False,
    }
    r = httpx.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60.0,
    )
    r.raise_for_status()
    log.info("Sent notification email to %s", to_addr)


def run_transcribe(project_root: Path, media: Path, txt_out: Path) -> None:
    script = project_root / "transcribe.py"
    py = project_root / ".venv" / "bin" / "python"
    if not py.is_file():
        py = Path(sys.executable)
    cmd = [str(py), str(script), str(media), "-o", str(txt_out)]
    log.info("Running: %s", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=str(project_root))


def process_one(
    token: str,
    item: dict,
    processed: set[str],
    *,
    state_path: Path,
    teams_transcript: bool,
    keep_local_copy: bool = False,
) -> None:
    item_id = item["id"]
    name = item["name"]
    stem = Path(name).stem
    drive_id = getenv("RECORDINGS_DRIVE_ID")
    prefix = "sp_teams_tr_" if teams_transcript else "sp_transcribe_"

    tmpdir = Path(tempfile.mkdtemp(prefix=prefix, dir=Path(__file__).resolve().parent / "tmp"))
    media_path = tmpdir / name
    txt_path = tmpdir / f"{stem}.txt"

    try:
        low = name.lower()
        if teams_transcript:
            allowed = transcript_suffixes()
            if not any(low.endswith(s) for s in allowed):
                log.warning(
                    "Skipping %s — not a Teams transcript extension %s (webhook may fire for all files)",
                    name,
                    allowed,
                )
                return

        log.info("Downloading %s", name)
        download_item(token, drive_id, item_id, media_path)
        log.info(
            "Saved with same filename as SharePoint: %s (full path: %s)",
            name,
            media_path,
        )

        if teams_transcript:
            low = name.lower()
            if low.endswith(".vtt"):
                raw = media_path.read_text(encoding="utf-8", errors="replace")
                txt_path.write_text(vtt_to_plain_text(raw), encoding="utf-8")
                log.info("Converted WebVTT → plain text: %s", txt_path.name)
            elif low.endswith(".txt"):
                shutil.copy2(media_path, txt_path)
                log.info("Using Teams .txt as transcript: %s", txt_path.name)
            else:
                raise RuntimeError(
                    f"Unsupported transcript type (set TEAMS_TRANSCRIPT_EXTENSIONS): {name!r}"
                )
        else:
            project_root = Path(getenv("PROJECT_ROOT")).resolve()
            run_transcribe(project_root, media_path, txt_path)

        txt_bytes = txt_path.read_bytes()
        s3_key = f"transcripts/{stem}.txt"
        s3_put_if_configured(txt_path, s3_key)

        tag = "TeamsTranscript" if teams_transcript else "Transcribe"
        try:
            upload_transcript(
                token,
                transcripts_drive_id(),
                transcripts_folder_id(),
                f"{stem}.txt",
                txt_bytes,
            )
            log.info("Uploaded transcript to SharePoint: %s.txt", stem)
        except Exception as e:
            log.exception("SharePoint transcript upload failed")
            try:
                send_mail_graph(
                    token,
                    subject=f"[{tag}] SharePoint upload failed: {stem}",
                    body_text=(
                        f"Source file: {name}\nError: {e!r}\n"
                        f"Transcript bytes: {len(txt_bytes)}\nS3 key (if configured): {s3_key}"
                    ),
                )
            except Exception as mail_e:
                log.error("Could not send mail: %s", mail_e)
            raise

        processed.add(item_id)
        save_processed(processed, state_path)

        if keep_local_copy:
            out_dir = Path(__file__).resolve().parent / "local_transcripts"
            out_dir.mkdir(parents=True, exist_ok=True)
            dest = out_dir / f"{stem}.txt"
            shutil.copy2(txt_path, dest)
            log.info("Saved local copy: %s", dest)
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Pull meeting media or Teams transcript files from SharePoint, produce .txt, upload."
    )
    ap.add_argument(
        "--list-files",
        action="store_true",
        dest="list_files",
        help="Print all files in the Recordings folder (name, size, Graph id) and exit.",
    )
    ap.add_argument(
        "--teams-transcript",
        action="store_true",
        dest="teams_transcript",
        help="Process Teams .vtt/.txt transcripts only (no Whisper). Uses processed_teams_transcripts.json.",
    )
    ap.add_argument(
        "--limit",
        type=int,
        default=None,
        metavar="N",
        help="Process at most N new recordings this run (e.g. 1 for a single test).",
    )
    ap.add_argument(
        "--skip",
        type=int,
        default=0,
        metavar="N",
        help="Skip the first N unprocessed recordings (e.g. --skip 1 --limit 1 for the second).",
    )
    ap.add_argument(
        "--item-id",
        dest="item_id",
        default=None,
        metavar="ID",
        help="Graph driveItem id of one file to process (recording or .vtt per mode).",
    )
    ap.add_argument(
        "--force",
        action="store_true",
        help="With --item-id, run even if that item is already marked processed.",
    )
    ap.add_argument(
        "--keep-local-copy",
        action="store_true",
        dest="keep_local_copy",
        help="Also save each .txt under automation/local_transcripts/.",
    )
    ap.add_argument(
        "--recursive",
        action="store_true",
        help="List/process files in subfolders under RECORDINGS_FOLDER_ID (see RECORDINGS_RECURSIVE).",
    )
    args = ap.parse_args()

    recurse = args.recursive or os.environ.get("RECORDINGS_RECURSIVE", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    if recurse:
        log.info(
            "Recursive folder scan enabled (nested .vtt/.txt or recordings under Recordings subfolders)"
        )

    if args.list_files:
        token = graph_token()
        rows = list_folder_all_files(token, recursive=recurse)
        log.info("%d file(s) under Recordings folder%s:", len(rows), " (recursive)" if recurse else "")
        for it in sorted(rows, key=lambda x: (x.get("name") or "").lower()):
            nm = it.get("name", "")
            sz = it.get("size")
            iid = it.get("id", "")
            print(f"{nm}\tsize={sz}\tid={iid}")
        return 0

    suffixes = transcript_suffixes() if args.teams_transcript else RECORDING_SUFFIXES
    state_path = STATE_PATH_TEAMS_TRANSCRIPTS if args.teams_transcript else STATE_PATH

    Path(__file__).resolve().parent.joinpath("tmp").mkdir(exist_ok=True)
    token = graph_token()
    processed = load_processed(state_path)
    drive_id = getenv("RECORDINGS_DRIVE_ID")

    if args.item_id:
        item = get_drive_item(token, drive_id, args.item_id)
        if "folder" in item:
            raise SystemExit("--item-id must be a file, not a folder")
        name = (item.get("name") or "").lower()
        if not any(name.endswith(s) for s in suffixes):
            raise SystemExit(
                f"Item name not allowed for this mode ({suffixes}): {item.get('name')}"
            )
        if item["id"] in processed and not args.force:
            log.info("Item already processed; pass --force to run again.")
            return 0
        new_items = [item]
    else:
        items = list_folder_files(token, suffixes, recursive=recurse)
        new_items = [it for it in items if it["id"] not in processed]
        if args.skip < 0:
            raise SystemExit("--skip must be >= 0")
        if args.skip:
            if args.skip >= len(new_items):
                log.info(
                    "Cannot skip %d: only %d unprocessed recording(s). Nothing to do.",
                    args.skip,
                    len(new_items),
                )
                new_items = []
            else:
                log.info(
                    "Skipping first %d unprocessed recording(s); continuing with the rest",
                    args.skip,
                )
                new_items = new_items[args.skip :]
        if args.limit is not None and args.limit >= 0:
            new_items = new_items[: args.limit]

    if not new_items:
        log.info("No new items to process")
        return 0
    mode = "Teams transcript (.vtt/.txt)" if args.teams_transcript else "Recording → Whisper"
    log.info("%d file(s) to process this run (%s)", len(new_items), mode)
    for it in new_items:
        try:
            process_one(
                token,
                it,
                processed,
                state_path=state_path,
                teams_transcript=args.teams_transcript,
                keep_local_copy=args.keep_local_copy,
            )
        except Exception:
            log.exception("Failed item %s — will retry on next run", it.get("name"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
