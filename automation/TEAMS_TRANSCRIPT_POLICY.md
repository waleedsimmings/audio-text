# Enable Automatic Transcript Saving for Future Teams Meetings

This configures Teams to automatically save `.vtt` transcript files alongside meeting recordings in SharePoint.

**Note:** This only affects FUTURE meetings. Old recordings will NOT retroactively get transcript files.

---

## Option 1: Teams Admin Center (GUI)

1. Go to [Teams Admin Center](https://admin.teams.microsoft.com)

2. Navigate to **Meetings** → **Meeting policies**

3. Select **Global (Org-wide default)** or create a custom policy

4. Scroll to **Recording & transcription** section

5. Enable:
   - **Transcription** = On
   - **Live captions** = On (optional)

6. Click **Save**

---

## Option 2: PowerShell Commands

Connect to Teams PowerShell (requires Teams Admin role):

```powershell
# Install Teams PowerShell module if not installed
Install-Module -Name MicrosoftTeams -Force -AllowClobber

# Connect (will prompt for credentials)
Connect-MicrosoftTeams

# Enable transcription globally
Set-CsTeamsMeetingPolicy -Identity Global -AllowTranscription $true

# Verify the setting
Get-CsTeamsMeetingPolicy -Identity Global | Select-Object AllowTranscription
```

---

## Configure Recording Expiration (Optional)

By default, recordings expire after 120 days. To change or disable:

```powershell
# Never expire recordings (set to -1)
Set-CsTeamsMeetingPolicy -Identity Global -NewMeetingRecordingExpirationDays -1

# Or set to 365 days
Set-CsTeamsMeetingPolicy -Identity Global -NewMeetingRecordingExpirationDays 365

# Verify
Get-CsTeamsMeetingPolicy -Identity Global | Select-Object NewMeetingRecordingExpirationDays
```

---

## Where Transcripts Are Saved

After enabling, transcripts will be saved:

| Meeting Type | Location |
|--------------|----------|
| Channel meeting | SharePoint site → Recordings folder (same as .mp4) |
| Private/ad-hoc meeting | Organizer's OneDrive → Recordings folder |

The transcript file will have the same name as the recording with `.vtt` extension.

---

## Verify It's Working

After a new meeting with recording:

1. Go to SharePoint/OneDrive Recordings folder
2. Look for a `.vtt` file next to the `.mp4`
3. If present, your existing `poll_sharepoint_transcribe.py --teams-transcript` script will work

---

## For Old Meetings

Old meetings already have transcripts embedded in Stream, but the policy won't extract them retroactively. For those, use one of:

1. **Certificate-based script**: `fetch_sharepoint_stream_transcripts.py` (see `SETUP_CERTIFICATE.md`)
2. **Manual download**: Open each video in SharePoint, click Transcript → Download as .vtt

---

## Additional Resources

- [Manage transcription and captions](https://learn.microsoft.com/en-us/microsoftteams/meeting-transcription-captions)
- [Teams meeting recording policies](https://learn.microsoft.com/en-us/microsoftteams/meeting-recording)
