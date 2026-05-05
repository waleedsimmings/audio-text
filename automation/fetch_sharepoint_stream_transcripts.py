#!/usr/bin/env python3
"""
Download transcripts from SharePoint Stream videos using the SharePoint v2.1 API.

This works for Teams meeting recordings stored in SharePoint where the transcript
exists in Stream (visible via "Download as .vtt" in the player) but is NOT saved
as a separate file in the SharePoint folder.

Uses undocumented SharePoint v2.1 API endpoint:
  https://{tenant}.sharepoint.com/_api/v2.1/drives/{driveId}/items/{fileId}/media/transcripts

IMPORTANT: SharePoint REST API requires CERTIFICATE-based authentication.
Client secrets DO NOT work (you'll get "401 Unsupported app only token").

Setup:
  1. Generate a self-signed certificate (see SETUP_CERTIFICATE.md)
  2. Upload the .cer file to your Azure AD app registration
  3. Set CERT_THUMBPRINT and CERT_PATH in .env

Requires Azure AD app with:
  - Sites.Read.All (SharePoint API scope, not just Graph)
  
Env (.env):
  TENANT_ID, CLIENT_ID
  CERT_THUMBPRINT       - Certificate thumbprint (from Azure AD or openssl)
  CERT_PATH             - Path to .pem or .pfx certificate file (with private key)
  CERT_PASSWORD         - Password for .pfx file (optional, leave empty for .pem)
  RECORDINGS_DRIVE_ID   - Graph drive ID for the Recordings folder
  RECORDINGS_FOLDER_ID  - Graph folder item ID
  SHAREPOINT_HOSTNAME   - e.g. fashco.sharepoint.com

Output: automation/local_transcripts_stream/<recording-name>.vtt
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

import httpx
from dotenv import load_dotenv
from msal import ConfidentialClientApplication
from cryptography.hazmat.primitives import serialization
from cryptography.hazmat.backends import default_backend
from cryptography import x509

load_dotenv(Path(__file__).resolve().parent / ".env")

GRAPH = "https://graph.microsoft.com/v1.0"
OUTPUT_DIR = Path(__file__).resolve().parent / "local_transcripts_stream"
STATE_PATH = Path(__file__).resolve().parent / "processed_stream_transcripts.json"


def getenv(name: str, default: str | None = None) -> str:
    v = os.environ.get(name, "").strip()
    if not v:
        if default is not None:
            return default
        raise SystemExit(f"Missing env var: {name}")
    return v


def load_certificate() -> dict:
    """
    Load certificate for MSAL authentication.
    
    Returns a dict suitable for ConfidentialClientApplication's client_credential parameter.
    Supports both .pem and .pfx formats.
    """
    cert_path = Path(getenv("CERT_PATH"))
    if not cert_path.is_file():
        raise SystemExit(f"Certificate file not found: {cert_path}")
    
    cert_password = os.environ.get("CERT_PASSWORD", "").strip()
    password_bytes = cert_password.encode() if cert_password else None
    
    thumbprint = getenv("CERT_THUMBPRINT")
    
    cert_data = cert_path.read_bytes()
    
    if cert_path.suffix.lower() == ".pfx":
        from cryptography.hazmat.primitives.serialization import pkcs12
        private_key, certificate, _ = pkcs12.load_key_and_certificates(
            cert_data, password_bytes, default_backend()
        )
        private_key_pem = private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption(),
        ).decode("utf-8")
        cert_pem = certificate.public_bytes(serialization.Encoding.PEM).decode("utf-8")
    else:
        private_key_pem = cert_data.decode("utf-8")
        cert_pem_match = None
        if "-----BEGIN CERTIFICATE-----" in private_key_pem:
            import re as regex
            cert_match = regex.search(
                r"(-----BEGIN CERTIFICATE-----.*?-----END CERTIFICATE-----)",
                private_key_pem,
                regex.DOTALL,
            )
            if cert_match:
                cert_pem = cert_match.group(1)
            else:
                cert_pem = ""
        else:
            cert_file = cert_path.with_suffix(".cer")
            if cert_file.is_file():
                cert_pem = cert_file.read_text()
            else:
                cert_pem = ""
    
    return {
        "private_key": private_key_pem,
        "thumbprint": thumbprint,
        "public_certificate": cert_pem if cert_pem else None,
    }


def graph_token() -> str:
    """
    Get a token for Microsoft Graph API.
    
    Uses client secret (if available) since it already has Graph permissions.
    Certificate auth may not have User.Read.All for speaker name resolution.
    """
    client_secret = os.environ.get("CLIENT_SECRET", "").strip()
    
    if client_secret:
        # Use client secret for Graph (already has permissions)
        app = ConfidentialClientApplication(
            getenv("CLIENT_ID"),
            authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
            client_credential=client_secret,
        )
    else:
        # Fall back to certificate
        cert_cred = load_certificate()
        app = ConfidentialClientApplication(
            getenv("CLIENT_ID"),
            authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
            client_credential=cert_cred,
        )
    
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise SystemExit(f"Graph token error: {result.get('error_description') or result}")
    return result["access_token"]


def sharepoint_token(hostname: str) -> str:
    """
    Get a token for SharePoint REST API using certificate auth.
    
    SharePoint REST API (_api) REQUIRES certificate-based authentication.
    Client secrets will fail with "401 Unsupported app only token".
    """
    cert_cred = load_certificate()
    resource = f"https://{hostname}"
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=cert_cred,
    )
    result = app.acquire_token_for_client(scopes=[f"{resource}/.default"])
    if "access_token" not in result:
        raise SystemExit(f"SharePoint token error: {result.get('error_description') or result}")
    return result["access_token"]


def safe_filename(s: str) -> str:
    return re.sub(r"[^\w.\-]+", "_", s)[:200] or "transcript"


# Cache for user display names (user_id -> display_name)
_user_cache: dict[str, str] = {}


def resolve_user_id(user_id: str, g_token: str) -> str:
    """
    Resolve Azure AD user ID to display name using Graph API.
    
    The VTT has speaker IDs like: 8625500f-2b8b-4bc5-a70c-9fef616f7822@ff82aff0-cc0d-4103-98d0-16f4c95075ad
    We extract the user ID (before @) and look it up.
    """
    if "@" in user_id:
        user_id = user_id.split("@")[0]
    
    if not user_id or len(user_id) < 10:
        return user_id
    
    if user_id in _user_cache:
        return _user_cache[user_id]
    
    try:
        url = f"{GRAPH}/users/{user_id}?$select=displayName"
        r = httpx.get(url, headers={"Authorization": f"Bearer {g_token}"}, timeout=30.0)
        if r.status_code == 200:
            name = r.json().get("displayName", "").strip()
            if name:
                _user_cache[user_id] = name
                return name
    except Exception:
        pass
    
    _user_cache[user_id] = user_id
    return user_id


def replace_speaker_ids_in_vtt(vtt_content: bytes, g_token: str) -> bytes:
    """
    Replace speaker IDs in VTT with actual display names.
    
    Converts: <v 8625500f-2b8b-4bc5-a70c-9fef616f7822@tenant>Text
    Or:       <v 8625500f-2b8b-4bc5-a70c-9fef616f7822>Text
    To:       <v Waleed Tahir>Text
    """
    text = vtt_content.decode("utf-8", errors="replace")
    
    # Find all unique speaker IDs - UUID format with optional @tenant suffix
    # Pattern matches: <v UUID> or <v UUID@tenant>
    pattern = r"<v\s+([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}(?:@[0-9a-fA-F\-]+)?)\s*>"
    matches = set(re.findall(pattern, text))
    
    if not matches:
        return vtt_content
    
    print(f"  Resolving {len(matches)} speaker ID(s)...")
    
    # Build replacement map
    replacements = {}
    unresolved = []
    for speaker_id in matches:
        # Extract just the user UUID (before @)
        user_uuid = speaker_id.split("@")[0] if "@" in speaker_id else speaker_id
        display_name = resolve_user_id(user_uuid, g_token)
        if display_name and display_name != user_uuid:
            replacements[speaker_id] = display_name
            print(f"    {user_uuid[:8]}... -> {display_name}")
        else:
            unresolved.append(speaker_id)
            # For unresolved (likely external guests), use "Guest" or short ID
            replacements[speaker_id] = f"Guest-{user_uuid[:4]}"
            print(f"    {user_uuid[:8]}... -> Guest-{user_uuid[:4]} (external user)")
    
    # Replace all occurrences
    for old_id, new_name in replacements.items():
        text = text.replace(f"<v {old_id}>", f"<v {new_name}>")
    
    return text.encode("utf-8")


def convert_vtt_to_plain_text(vtt_content: bytes) -> bytes:
    """
    Convert VTT to clean plain text format.
    
    Removes:
    - WEBVTT header
    - Cue numbers (1, 2, 3...)
    - Timestamps (00:00:00.000 --> 00:00:01.000)
    
    Converts: <v Speaker Name>Text
    To:       Speaker Name: Text
    """
    text = vtt_content.decode("utf-8", errors="replace")
    lines = text.splitlines()
    
    output_lines = []
    prev_speaker = None
    prev_text_parts = []
    
    for line in lines:
        line = line.strip()
        
        # Skip empty lines, WEBVTT header, cue numbers, timestamps
        if not line:
            continue
        if line.upper().startswith("WEBVTT"):
            continue
        if line.isdigit():
            continue
        if "-->" in line:
            continue
        if line.lower().startswith("kind:") or line.lower().startswith("language:"):
            continue
        
        # Parse speaker and text from <v Speaker>Text format
        match = re.match(r"<v\s+([^>]+)>(.*)$", line)
        if match:
            speaker = match.group(1).strip()
            content = match.group(2).strip()
            
            # Combine consecutive lines from same speaker
            if speaker == prev_speaker:
                prev_text_parts.append(content)
            else:
                # Output previous speaker's combined text
                if prev_speaker and prev_text_parts:
                    combined = " ".join(prev_text_parts)
                    output_lines.append(f"{prev_speaker}: {combined}")
                
                prev_speaker = speaker
                prev_text_parts = [content] if content else []
        else:
            # Line without speaker tag - append to previous
            if prev_text_parts is not None:
                prev_text_parts.append(line)
    
    # Don't forget the last speaker
    if prev_speaker and prev_text_parts:
        combined = " ".join(prev_text_parts)
        output_lines.append(f"{prev_speaker}: {combined}")
    
    return "\n\n".join(output_lines).encode("utf-8")


def load_processed() -> set[str]:
    if not STATE_PATH.is_file():
        return set()
    data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    return set(data.get("file_ids", []))


def save_processed(ids: set[str]) -> None:
    STATE_PATH.write_text(
        json.dumps({"file_ids": sorted(ids)}, indent=2),
        encoding="utf-8",
    )


def list_recordings(token: str, drive_id: str, folder_id: str) -> list[dict]:
    """List all .mp4 files in the Recordings folder (recursive)."""
    items: list[dict] = []
    
    def walk(fid: str, depth: int = 0) -> None:
        if depth > 10:
            return
        url = f"{GRAPH}/drives/{drive_id}/items/{fid}/children"
        while url:
            r = httpx.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=120.0)
            r.raise_for_status()
            body = r.json()
            for it in body.get("value", []):
                if "folder" in it:
                    walk(it["id"], depth + 1)
                else:
                    name = (it.get("name") or "").lower()
                    if name.endswith(".mp4"):
                        items.append(it)
            url = body.get("@odata.nextLink")
    
    walk(folder_id)
    return items


def get_transcript_via_sharepoint_api(
    sp_token: str,
    hostname: str,
    drive_id: str,
    file_id: str,
) -> bytes | None:
    """
    Fetch transcript using SharePoint v2.1 API.
    
    Endpoint: /_api/v2.1/drives/{driveId}/items/{fileId}/media/transcripts
    Returns the transcript content or None if not available.
    """
    url = f"https://{hostname}/_api/v2.1/drives/{drive_id}/items/{file_id}/media/transcripts"
    headers = {
        "Authorization": f"Bearer {sp_token}",
        "Accept": "application/json",
    }
    
    r = httpx.get(url, headers=headers, timeout=60.0)
    
    if r.status_code == 404:
        return None
    if r.status_code >= 400:
        print(f"  Transcript API error {r.status_code}: {r.text[:500]}", file=sys.stderr)
        return None
    
    try:
        body = r.json()
    except Exception as e:
        print(f"  Failed to parse transcript response: {e}", file=sys.stderr)
        return None
    
    value = body.get("value") or []
    if not value:
        return None
    
    download_url = None
    for v in value:
        download_url = v.get("temporaryDownloadUrl") or v.get("@content.downloadUrl")
        if download_url:
            break
    
    if not download_url:
        stream_url = None
        for v in value:
            tid = v.get("id")
            if tid:
                stream_url = f"https://{hostname}/_api/v2.1/drives/{drive_id}/items/{file_id}/media/transcripts/{tid}/streamContent?format=vtt"
                break
        if stream_url:
            download_url = stream_url
    
    if not download_url:
        print(f"  No download URL in transcript response: {json.dumps(body)[:300]}", file=sys.stderr)
        return None
    
    dr = httpx.get(download_url, headers=headers, timeout=120.0, follow_redirects=True)
    if dr.status_code >= 400:
        print(f"  Transcript download error {dr.status_code}: {dr.text[:300]}", file=sys.stderr)
        return None
    
    return dr.content


def get_transcript_via_stream_content(
    sp_token: str,
    hostname: str,
    drive_id: str,
    file_id: str,
) -> bytes | None:
    """
    Alternative: Try streamContent endpoint directly.
    
    Some community posts suggest:
    /_api/v2.1/drives/{driveId}/items/{fileId}/media/transcripts/{id}/streamContent?format=vtt
    
    But we need the transcript ID first, so we list transcripts then fetch.
    """
    url = f"https://{hostname}/_api/v2.1/drives/{drive_id}/items/{file_id}/media/transcripts"
    headers = {
        "Authorization": f"Bearer {sp_token}",
        "Accept": "application/json",
    }
    
    r = httpx.get(url, headers=headers, timeout=60.0)
    if r.status_code >= 400:
        return None
    
    try:
        body = r.json()
        value = body.get("value") or []
    except Exception:
        return None
    
    for v in value:
        tid = v.get("id")
        if not tid:
            continue
        
        for fmt in ("vtt", "json"):
            stream_url = f"https://{hostname}/_api/v2.1/drives/{drive_id}/items/{file_id}/media/transcripts/{tid}/streamContent?format={fmt}"
            sr = httpx.get(stream_url, headers=headers, timeout=120.0, follow_redirects=True)
            if sr.status_code < 400 and len(sr.content) > 50:
                if fmt == "json":
                    return convert_json_transcript_to_vtt(sr.content)
                return sr.content
    
    return None


def convert_json_transcript_to_vtt(data: bytes) -> bytes:
    """Convert JSON transcript format to VTT."""
    try:
        j = json.loads(data)
    except Exception:
        return data
    
    lines = ["WEBVTT", ""]
    
    entries = j if isinstance(j, list) else j.get("cues") or j.get("entries") or []
    
    for i, entry in enumerate(entries, 1):
        start = entry.get("start") or entry.get("startTime") or 0
        end = entry.get("end") or entry.get("endTime") or start + 1
        text = entry.get("text") or entry.get("content") or ""
        speaker = entry.get("speaker") or entry.get("speakerId") or ""
        
        def ms_to_vtt(ms: float) -> str:
            if isinstance(ms, str):
                return ms
            total_sec = ms / 1000.0 if ms > 1000 else ms
            h = int(total_sec // 3600)
            m = int((total_sec % 3600) // 60)
            s = total_sec % 60
            return f"{h:02d}:{m:02d}:{s:06.3f}"
        
        lines.append(str(i))
        lines.append(f"{ms_to_vtt(start)} --> {ms_to_vtt(end)}")
        if speaker:
            lines.append(f"<v {speaker}>{text}")
        else:
            lines.append(text)
        lines.append("")
    
    return "\n".join(lines).encode("utf-8")


def upload_to_sharepoint(
    g_token: str,
    drive_id: str,
    folder_id: str,
    filename: str,
    content: bytes,
) -> bool:
    """
    Upload a file to SharePoint using Graph API.
    
    Uses PUT to /drives/{driveId}/items/{folderId}:/{filename}:/content
    """
    from urllib.parse import quote
    
    encoded_name = quote(filename, safe="")
    url = f"{GRAPH}/drives/{drive_id}/items/{folder_id}:/{encoded_name}:/content"
    
    headers = {
        "Authorization": f"Bearer {g_token}",
        "Content-Type": "text/plain; charset=utf-8",
    }
    
    try:
        r = httpx.put(url, headers=headers, content=content, timeout=120.0)
        if r.status_code in (200, 201):
            return True
        print(f"  Upload error {r.status_code}: {r.text[:200]}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"  Upload exception: {e}", file=sys.stderr)
        return False


def main() -> int:
    import argparse
    
    ap = argparse.ArgumentParser(description="Download Stream video transcripts from SharePoint")
    ap.add_argument("--list", action="store_true", help="List recordings and their transcript status")
    ap.add_argument("--limit", type=int, default=None, help="Process at most N files")
    ap.add_argument("--force", action="store_true", help="Re-download even if already processed")
    ap.add_argument("--file-id", dest="file_id", help="Process only this specific file ID")
    ap.add_argument("--plain-text", dest="plain_text", action="store_true", 
                    help="Output clean plain text (.txt) instead of VTT format")
    ap.add_argument("--upload", action="store_true",
                    help="Upload transcript to SharePoint Recordings folder")
    ap.add_argument(
        "--no-state",
        action="store_true",
        dest="no_state",
        help="Do not read/write processed_stream_transcripts.json (for webhooks / retries)",
    )
    args = ap.parse_args()
    
    hostname = getenv("SHAREPOINT_HOSTNAME", "fashco.sharepoint.com")
    drive_id = getenv("RECORDINGS_DRIVE_ID")
    folder_id = getenv("RECORDINGS_FOLDER_ID")
    
    print(f"SharePoint: {hostname}")
    print(f"Drive ID: {drive_id}")
    print(f"Folder ID: {folder_id}")
    print()
    
    print("Acquiring tokens...")
    g_token = graph_token()
    sp_token = sharepoint_token(hostname)
    print("OK: Got Graph and SharePoint tokens")
    print()
    
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    processed = set() if args.no_state else load_processed()
    
    if args.file_id:
        url = f"{GRAPH}/drives/{drive_id}/items/{args.file_id}"
        r = httpx.get(url, headers={"Authorization": f"Bearer {g_token}"}, timeout=60.0)
        r.raise_for_status()
        item = r.json()
        if "folder" in item:
            print("Skipping folder item", file=sys.stderr)
            return 0
        nm = (item.get("name") or "").lower()
        if not nm.endswith(".mp4"):
            print(f"Skipping non-recording: {item.get('name')}", file=sys.stderr)
            return 0
        recordings = [item]
    else:
        print("Listing recordings...")
        recordings = list_recordings(g_token, drive_id, folder_id)
        print(f"Found {len(recordings)} .mp4 files")
        print()
    
    if args.list:
        for rec in recordings:
            name = rec.get("name", "?")
            fid = rec.get("id", "?")
            status = "processed" if fid in processed else "pending"
            print(f"  [{status}] {name}")
            print(f"           id={fid}")
        return 0
    
    to_process = []
    for rec in recordings:
        fid = rec.get("id")
        if fid in processed and not args.force:
            continue
        to_process.append(rec)
    
    if args.limit:
        to_process = to_process[:args.limit]
    
    if not to_process:
        print("No new recordings to process")
        return 0
    
    print(f"Processing {len(to_process)} recording(s)...")
    print()
    
    success = 0
    failed = 0
    no_transcript = 0
    
    for rec in to_process:
        name = rec.get("name", "unknown")
        fid = rec.get("id")
        print(f"Processing: {name}")
        
        content = get_transcript_via_sharepoint_api(sp_token, hostname, drive_id, fid)
        
        if not content:
            content = get_transcript_via_stream_content(sp_token, hostname, drive_id, fid)
        
        if not content:
            print(f"  No transcript available")
            no_transcript += 1
            if not args.no_state:
                processed.add(fid)
                save_processed(processed)
            continue
        
        # Replace speaker IDs with display names
        content = replace_speaker_ids_in_vtt(content, g_token)
        
        # Convert to plain text if requested
        if args.plain_text:
            content = convert_vtt_to_plain_text(content)
            ext = ".txt"
        else:
            ext = ".vtt"
        
        stem = Path(name).stem
        out_name = safe_filename(stem) + ext
        out_path = OUTPUT_DIR / out_name
        out_path.write_bytes(content)
        print(f"  Saved locally: {out_path} ({len(content)} bytes)")
        
        # Upload to SharePoint if requested
        if args.upload:
            # Use original filename stem + .txt for SharePoint
            sp_filename = stem + ext
            if upload_to_sharepoint(g_token, drive_id, folder_id, sp_filename, content):
                print(f"  Uploaded to SharePoint: {sp_filename}")
            else:
                print(f"  Failed to upload to SharePoint", file=sys.stderr)
        
        if not args.no_state:
            processed.add(fid)
            save_processed(processed)
        success += 1
    
    print()
    print(f"Done: {success} downloaded, {no_transcript} no transcript, {failed} failed")
    
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
