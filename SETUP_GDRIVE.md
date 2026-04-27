# Google Drive Upload — One-Time Setup

The app uploads categorized invoices to Google Drive using **OAuth user delegation**.
You (the developer) do this setup **once**. End-clients of the app never paste any
credentials.

Destination folder is fixed at:
`https://drive.google.com/drive/folders/1-ivSbgBMGZJwqZ2hp3r7jH4HkzEXjWe7`

## You don't need the client's Google password

You sign in with **your own** Google account. The client just needs to share the
folder with your email as **Editor**, one time. After that, your account can
create files inside that folder forever — uploaded files will live in the
client's folder structure, owned by your account, counting against your Drive
storage (15 GB free).

**Step 0 (ask the client):** open the folder in Drive → **Share** → add your
email (e.g. `firas.latrach@horizon-tech.tn`) with the **Editor** role → **Send**.

That's the only thing the client ever does.

## Why not service accounts?

The previous setup used a service account JSON and failed with
`storageQuotaExceeded`. Service accounts have **no personal Drive storage** — they
can only own files inside Google Workspace **Shared Drives**. Switching to OAuth
user creds means uploads land in the client's folder using your account's storage.

## One-time setup (10 minutes)

### 1. Enable the Drive API

[Google Cloud Console](https://console.cloud.google.com/) → select or create a
project → **APIs & Services → Library** → search **Google Drive API** → **Enable**.

### 2. Configure the OAuth consent screen

**APIs & Services → OAuth consent screen**:

- User type: **External**
- App name: anything (e.g. "Invoice Categorizer")
- User support email: your email
- Developer contact: your email
- Scopes: skip (we request scopes from the script)
- **Test users:** add **your own email** (the one that has Editor access to the
  client's folder — *not* the client's email).

Leave the app in **Testing** mode. **Important:** in Testing mode, refresh tokens
expire after 7 days *unless* you are listed as a Test user. Adding yourself
prevents the expiry. (If you ever need it to work for other accounts, click
**Publish app**, but for this single-user case, Testing + test user is fine.)

### 3. Create the OAuth client (Desktop app type)

**APIs & Services → Credentials → Create Credentials → OAuth client ID**:

- Application type: **Desktop app**
- Name: anything

Click **Download JSON** and save the file as:

```
scripts/oauth_client.json
```

(This file is a secret — already covered by `.gitignore`. Do not commit it.)

### 4. Run the helper script

```bash
pip install -r requirements.txt
python scripts/get_refresh_token.py
```

A browser window opens. Sign in with **your own Google account** (the one the
client shared the folder with in Step 0). Click "Allow". The script prints
something like:

```
[gdrive]
client_id = "1234....apps.googleusercontent.com"
client_secret = "GOCSPX-..."
refresh_token = "1//0g..."
folder_id = "1-ivSbgBMGZJwqZ2hp3r7jH4HkzEXjWe7"
```

### 5. Save the secrets

**Local development:** create `.streamlit/secrets.toml` and paste the block above
(template at `.streamlit/secrets.toml.example`).

**Streamlit Cloud:** App → **Settings → Secrets** → paste the same TOML block.

Done. `streamlit run app.py` → click **Upload to Google Drive** in the app — it
just works, every time, for every client.

## Troubleshooting

**"Google did not return a refresh_token"** — happens if you previously authorized
the same OAuth client. Revoke at <https://myaccount.google.com/permissions> and
re-run the helper.

**Refresh token stops working after 7 days** — your account is not on the OAuth
consent screen's Test users list. Add it (step 2) and re-run the helper.

**Upload fails with `invalid_grant`** — refresh token was revoked (manually, or
because the password changed, or 7-day expiry). Re-run the helper to generate a
new one.

**Upload fails with "File not found" or 404 on the folder ID** — your account
doesn't have access to the client's folder. Ask the client to re-share it with
your email as **Editor** (Step 0).
