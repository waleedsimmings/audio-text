# Certificate Setup for SharePoint REST API Access

SharePoint REST API (`/_api/`) requires **certificate-based authentication**. Client secrets do NOT work and will return "401 Unsupported app only token".

This guide walks you through creating a certificate and configuring your Azure AD app.

---

## Step 1: Generate a Self-Signed Certificate

### Option A: Using OpenSSL (macOS/Linux)

```bash
# Navigate to the automation folder
cd /Users/waleedtahir/Documents/audio-text/automation

# Generate a private key and certificate (valid for 2 years)
openssl req -x509 -newkey rsa:2048 -keyout sp_private.pem -out sp_cert.cer \
  -days 730 -nodes -subj "/CN=MeetingTranscribeApp"

# Combine into a single PEM file (key + cert)
cat sp_private.pem sp_cert.cer > sp_combined.pem

# Get the certificate thumbprint (you'll need this for .env)
openssl x509 -in sp_cert.cer -noout -fingerprint -sha1 | sed 's/://g' | cut -d= -f2
```

This creates:
- `sp_private.pem` - Private key (keep secure, never share)
- `sp_cert.cer` - Public certificate (upload to Azure AD)
- `sp_combined.pem` - Combined file for the script

### Option B: Using PowerShell (Windows)

```powershell
# Generate certificate
$cert = New-SelfSignedCertificate -DnsName "fashco.sharepoint.com" `
  -CertStoreLocation "cert:\CurrentUser\My" `
  -NotAfter (Get-Date).AddYears(2) `
  -KeyExportPolicy Exportable -KeySpec Signature

# Export public certificate (.cer) for Azure AD
Export-Certificate -Cert $cert -FilePath ".\sp_cert.cer"

# Export private key (.pfx) for the script
$password = ConvertTo-SecureString -String "YourPassword123" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath ".\sp_cert.pfx" -Password $password

# Get thumbprint
$cert.Thumbprint
```

---

## Step 2: Upload Certificate to Azure AD

1. Go to [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **App registrations**

2. Find your app (the one with CLIENT_ID from your .env)

3. Click **Certificates & secrets** in the left menu

4. Click the **Certificates** tab

5. Click **Upload certificate**

6. Select your `sp_cert.cer` file

7. Click **Add**

8. Note the **Thumbprint** shown after upload (should match what you generated)

---

## Step 3: Add SharePoint API Permission

Your app needs the SharePoint `Sites.Read.All` permission (not just Graph):

1. In your app registration, go to **API permissions**

2. Click **Add a permission**

3. Select **SharePoint** (not Microsoft Graph)

4. Choose **Application permissions**

5. Check **Sites.Read.All**

6. Click **Add permissions**

7. Click **Grant admin consent for [Your Tenant]**

---

## Step 4: Update .env

Add these new variables to your `.env` file:

```bash
# Certificate authentication for SharePoint REST API
CERT_PATH=/Users/waleedtahir/Documents/audio-text/automation/sp_combined.pem
CERT_THUMBPRINT=ABC123DEF456...  # The thumbprint from Step 1 or Azure AD

# If using .pfx instead of .pem:
# CERT_PATH=/Users/waleedtahir/Documents/audio-text/automation/sp_cert.pfx
# CERT_PASSWORD=YourPassword123
```

---

## Step 5: Test the Script

```bash
cd /Users/waleedtahir/Documents/audio-text
source .venv/bin/activate
python automation/fetch_sharepoint_stream_transcripts.py --list
```

If successful, you'll see a list of recordings. Then run:

```bash
python automation/fetch_sharepoint_stream_transcripts.py --limit 1
```

This should download the first transcript to `local_transcripts_stream/`.

---

## Troubleshooting

### "401 Unsupported app only token"
- You're still using client secret. Make sure CERT_PATH and CERT_THUMBPRINT are set.

### "AADSTS700027: Client assertion contains an invalid signature"
- Thumbprint mismatch. Verify the thumbprint in .env matches what's in Azure AD.

### "AADSTS7000215: Invalid client secret"  
- The certificate wasn't loaded correctly. Check CERT_PATH points to a valid file.

### "403 Access denied"
- SharePoint Sites.Read.All permission not granted, or admin consent not given.

---

## Security Notes

- **Never commit** `sp_private.pem`, `sp_combined.pem`, or `.pfx` files to git
- Add them to `.gitignore`:
  ```
  *.pem
  *.pfx
  *.cer
  ```
- Store certificates securely; they grant access to your SharePoint
- Certificates expire after the validity period (2 years in this example)
