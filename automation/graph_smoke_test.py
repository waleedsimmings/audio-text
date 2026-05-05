#!/usr/bin/env python3
"""Verify MeetingTranscribe app token + Graph access (same creds as poll_sharepoint_transcribe.py)."""

from __future__ import annotations

import os
import sys
from pathlib import Path
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv(Path(__file__).resolve().parent / ".env")

GRAPH = "https://graph.microsoft.com/v1.0"


def main() -> int:
    for k in ("TENANT_ID", "CLIENT_ID", "CLIENT_SECRET"):
        if not os.environ.get(k, "").strip():
            print(f"Missing {k} in automation/.env", file=sys.stderr)
            return 1

    app = ConfidentialClientApplication(
        os.environ["CLIENT_ID"].strip(),
        authority=f"https://login.microsoftonline.com/{os.environ['TENANT_ID'].strip()}",
        client_credential=os.environ["CLIENT_SECRET"].strip(),
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        print("Token error:", result.get("error_description") or result, file=sys.stderr)
        return 1
    token = result["access_token"]
    print("OK: acquired application token\n")

    # Site key: hostname,siteCollectionId,siteId — override with SHAREPOINT_SITE_ID in .env
    site_spec = os.environ.get(
        "SHAREPOINT_SITE_ID",
        "fashco.sharepoint.com,f52eb9c4-ba08-451b-adc9-7ec4a744a613,bea79fe2-e984-4f16-9026-435738eace71",
    ).strip()

    tests = [
        ("List a few sites (tenant scope)", "GET", f"{GRAPH}/sites?$top=3"),
        ("SharePoint site metadata", "GET", f"{GRAPH}/sites/{site_spec}"),
        (
            "Default document library (drive) — use MeetingTranscribe, NOT Graph Explorer",
            "GET",
            f"{GRAPH}/sites/{site_spec}/drive",
        ),
    ]
    drive_id = os.environ.get("RECORDINGS_DRIVE_ID", "").strip()
    if drive_id:
        tests.append(("Drive metadata by RECORDINGS_DRIVE_ID", "GET", f"{GRAPH}/drives/{drive_id}"))
        folder_path = os.environ.get(
            "RECORDINGS_FOLDER_PATH", "Newcleux Product and Tech/Recordings"
        ).strip()
        enc = quote(folder_path, safe="/")
        tests.append(
            (
                f"Recordings folder ({folder_path})",
                "GET",
                f"{GRAPH}/drives/{drive_id}/root:/{enc}:",
            )
        )

    for label, method, url in tests:
        r = httpx.request(
            method,
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=60.0,
        )
        print(f"--- {label} ---")
        print(f"{r.status_code} {url}")
        if r.status_code >= 400:
            print(r.text[:2000])
        else:
            body = r.json()
            if "value" in body:
                for it in body.get("value", [])[:5]:
                    print(" ", it.get("displayName") or it.get("name"), it.get("id", "")[:60])
            else:
                print(" ", "id:", body.get("id", "")[:80], "| name:", body.get("name", ""))
                if "driveType" in body:
                    print("  → put this in .env as RECORDINGS_DRIVE_ID=" + body.get("id", ""))
                if body.get("folder") is not None and body.get("name"):
                    print("  → put this in .env as RECORDINGS_FOLDER_ID=" + body.get("id", ""))
        print()

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
