#!/usr/bin/env python3
"""
Download Teams meeting transcripts via Microsoft Graph (callTranscript), NOT from SharePoint folder listing.

Use when Stream shows "Download as .vtt" but no .vtt/.txt appears next to the .mp4 in the library.

Requires Azure AD application permission (admin consent):
  OnlineMeetingTranscript.Read.All   (application)

Optional for filenames that include the real meeting title (GET .../onlineMeetings/{id}?$select=subject):
  OnlineMeetings.Read.All or OnlineMeetingArtifact.Read.All — tenant-dependent; without it the
  name falls back to \"Meeting\" while date still uses transcript timestamps above.

Teams admin (required for app-only): an **application access policy** must allow this
application to read online meetings / transcripts **on behalf of** the user id you pass
(TRANSCRIPT_ORGANIZER_USER_ID — typically the meeting organizer). Without it, listing may
partially work but **GET .../transcripts/.../content often returns 403**.

  https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy

Example (Teams PowerShell, Global admin or Teams admin):

  New-CsApplicationAccessPolicy -Identity TranscriptGraphApp -AppIds "<CLIENT_ID>" `
    -Description "Graph meeting transcripts"
  Grant-CsApplicationAccessPolicy -PolicyName TranscriptGraphApp -Identity "<organizer@tenant.com>"

Propagation can take up to 24–48 hours.

Docs:
  https://learn.microsoft.com/en-us/graph/api/onlinemeeting-getalltranscripts
  https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/meeting-transcripts/overview-transcripts

IMPORTANT: Microsoft states getAllTranscripts does NOT support transcripts from *channel meetings*
  (meetings rooted in a Teams channel / saved to SharePoint site Recordings). If your meetings are
  channel meetings, this API may return nothing — then you need tenant policy to save transcript
  files to the library, or use export/Compliance paths.

Env (.env):
  TENANT_ID, CLIENT_ID, CLIENT_SECRET  (same app or a dedicated one)
  TRANSCRIPT_ORGANIZER_USER_ID         Azure AD object id of the meeting organizer to filter on

Optional:
  TRANSCRIPT_START_ISO=2026-03-01T00:00:00Z
  TRANSCRIPT_END_ISO=2026-04-30T23:59:59Z
  TRANSCRIPT_VERBOSE_SKIPS=1       Full Graph JSON per skipped 400/404/410 (default: one line)
  TRANSCRIPT_SKIP_MEETING_SUBJECT=1  Filename uses date + Meeting only (no extra GET for subject)

Output: automation/local_transcripts_from_graph/<YYYY-MM-DD>_<meeting-subject>_<id-suffix>.vtt
  (subject from Graph onlineMeeting when available; TRANSCRIPT_SKIP_MEETING_SUBJECT=1 uses "Meeting" only)
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path
from urllib.parse import quote

import httpx
from dotenv import load_dotenv
from msal import ConfidentialClientApplication

load_dotenv(Path(__file__).resolve().parent / ".env")

GRAPH = "https://graph.microsoft.com/v1.0"


def getenv(name: str) -> str:
    v = os.environ.get(name, "").strip()
    if not v:
        raise SystemExit(f"Missing env var: {name}")
    return v


def token() -> str:
    app = ConfidentialClientApplication(
        getenv("CLIENT_ID"),
        authority=f"https://login.microsoftonline.com/{getenv('TENANT_ID')}",
        client_credential=getenv("CLIENT_SECRET"),
    )
    r = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in r:
        raise SystemExit(r.get("error_description") or str(r))
    return r["access_token"]


def safe_filename(s: str) -> str:
    return re.sub(r"[^\w.\-]+", "_", s)[:180] or "transcript"


def graph_error_message(body_text: str) -> str:
    try:
        data = json.loads(body_text)
        err = data.get("error") or {}
        return str(err.get("message") or err.get("code") or body_text)[:240]
    except Exception:
        return body_text[:240]


_SKIP_TRANSCRIPT_HTTP = frozenset({400, 404, 410})


def meeting_subject(
    org: str,
    meeting_id: str,
    headers: dict[str, str],
    cache: dict[str, str],
) -> str:
    if meeting_id in cache:
        return cache[meeting_id]
    enc = quote(meeting_id, safe="")
    url = f"{GRAPH}/users/{org}/onlineMeetings/{enc}?$select=subject"
    r = httpx.get(url, headers=headers, timeout=60.0)
    if r.status_code != 200:
        cache[meeting_id] = ""
        return ""
    try:
        subj = (r.json().get("subject") or "").strip()
    except Exception:
        subj = ""
    cache[meeting_id] = subj
    return subj


def transcript_date_label(item: dict) -> str:
    """Best calendar day from Graph transcript row; ignore Teams placeholder 0001-01-01 dates."""
    for key in ("createdDateTime", "startDateTime", "endDateTime"):
        raw = (item.get(key) or "").strip()
        if len(raw) < 10:
            continue
        day = raw[:10]
        if day.startswith("0001-01-01") or day.startswith("0000-"):
            continue
        if day < "1990-01-01":
            continue
        return day
    return "undated"


def vtt_output_name(
    org: str,
    item: dict,
    tid: str,
    headers: dict[str, str],
    subject_cache: dict[str, str],
) -> str:
    date_part = transcript_date_label(item)
    meeting_id = (item.get("meetingId") or "").strip()
    skip_subject = os.environ.get("TRANSCRIPT_SKIP_MEETING_SUBJECT", "").strip().lower() in (
        "1",
        "true",
        "yes",
    )
    if skip_subject or not meeting_id:
        subj = "Meeting"
    else:
        subj = meeting_subject(org, meeting_id, headers, subject_cache) or "Meeting"
    slug = safe_filename(f"{date_part}_{subj}")
    slug = slug[:130] if len(slug) > 130 else slug
    tail = safe_filename(tid)
    tail = tail[-28:] if len(tail) > 28 else tail
    return f"{slug}_{tail}.vtt"


def main() -> int:
    org = getenv("TRANSCRIPT_ORGANIZER_USER_ID")
    t = token()
    headers = {"Authorization": f"Bearer {t}"}

    start = os.environ.get("TRANSCRIPT_START_ISO", "").strip()
    end = os.environ.get("TRANSCRIPT_END_ISO", "").strip()
    if start and end:
        url = (
            f"{GRAPH}/users/{org}/onlineMeetings/getAllTranscripts"
            f"(meetingOrganizerUserId='{org}',startDateTime={start},endDateTime={end})"
        )
    else:
        filt = f"$filter=MeetingOrganizer/User/Id eq '{org}'"
        url = f"{GRAPH}/users/{org}/onlineMeetings/getAllTranscripts?{filt}"

    out_dir = Path(__file__).resolve().parent / "local_transcripts_from_graph"
    out_dir.mkdir(parents=True, exist_ok=True)

    count = 0
    skipped = 0
    subject_cache: dict[str, str] = {}
    while url:
        r = httpx.get(url, headers=headers, timeout=120.0)
        if r.status_code >= 400:
            print(r.text[:4000], file=sys.stderr)
            r.raise_for_status()
        body = r.json()
        for item in body.get("value", []):
            tid = item.get("id") or "unknown"
            meeting_id = item.get("meetingId")
            api_content = (item.get("transcriptContentUrl") or "").strip()

            candidates = []
            if api_content:
                candidates.append(api_content)
            if meeting_id and tid != "unknown":
                built = (
                    f"{GRAPH}/users/{org}/onlineMeetings/{meeting_id}"
                    f"/transcripts/{tid}/content"
                )
                candidates.append(f"{built}?$format=text/vtt")
                candidates.append(built)

            seen: set[str] = set()
            cr = None
            for u in candidates:
                if not u or u in seen:
                    continue
                seen.add(u)
                cr = httpx.get(u, headers=headers, timeout=120.0)
                if cr.status_code < 400:
                    break

            if cr is None:
                continue

            if cr.status_code < 400:
                fn = vtt_output_name(org, item, tid, headers, subject_cache)
                path = out_dir / fn
                path.write_bytes(cr.content)
                print("Wrote", path, "bytes=", len(cr.content))
                count += 1
            elif cr.status_code in _SKIP_TRANSCRIPT_HTTP:
                verbose = os.environ.get("TRANSCRIPT_VERBOSE_SKIPS", "").strip().lower() in (
                    "1",
                    "true",
                    "yes",
                )
                if verbose:
                    print(
                        f"Transcript content HTTP {cr.status_code} (tried {len(seen)} URL(s)). "
                        "Graph body:\n"
                        + cr.text[:4000],
                        file=sys.stderr,
                    )
                    if cr.status_code == 400:
                        print(
                            "400 often means expired/not finalized/unsupported meeting type — "
                            "see script docstring.",
                            file=sys.stderr,
                        )
                tid_short = (tid[:56] + "…") if len(tid) > 56 else tid
                print(
                    f"Skip HTTP {cr.status_code} transcript id={tid_short!r} — {graph_error_message(cr.text)}",
                    file=sys.stderr,
                )
                skipped += 1
            else:
                if cr.status_code == 403:
                    body_snip = cr.text[:4000]
                    specific = ""
                    if "Application is not allowed to perform operations on the user" in body_snip:
                        specific = (
                            f"\nGraph is refusing access for the user in the URL path (should match "
                            f"TRANSCRIPT_ORGANIZER_USER_ID, currently {org!r}). Create the policy with "
                            f"-AppIds exactly this app's CLIENT_ID, then Grant-CsApplicationAccessPolicy "
                            f"to that user's sign-in name (UPN). 'RSC permission evaluation' in the "
                            f"message is normal for app-only; it does not replace the Teams policy.\n"
                        )
                    print(
                        "403 Forbidden on transcript content. For client-credentials (app-only) access, "
                        "Microsoft requires a Teams Application Access Policy that ties your CLIENT_ID "
                        "to the organizer (the user in TRANSCRIPT_ORGANIZER_USER_ID). See script "
                        "docstring and:\n"
                        "https://learn.microsoft.com/en-us/graph/cloud-communication-online-meeting-application-access-policy"
                        + specific
                        + "\nGraph error body:\n"
                        + body_snip,
                        file=sys.stderr,
                    )
                print(
                    f"Transcript content HTTP {cr.status_code} (tried {len(seen)} URL variant(s)). "
                    "Graph body:\n"
                    + cr.text[:4000],
                    file=sys.stderr,
                )
                cr.raise_for_status()
        url = body.get("@odata.nextLink")

    if not count:
        print(
            "No transcripts returned. If these are *channel* meetings, Graph getAllTranscripts "
            "may not apply — see script docstring.",
            file=sys.stderr,
        )
    elif skipped:
        print(
            f"Skipped {skipped} transcript(s) (400/404/410). "
            "Set TRANSCRIPT_VERBOSE_SKIPS=1 for full Graph bodies.",
            file=sys.stderr,
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
