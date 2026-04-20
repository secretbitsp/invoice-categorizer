"""
Download PDF invoices from Google Drive or OneDrive folders.
"""

import base64
import io
import requests
from urllib.parse import parse_qs, urlencode, urljoin, urlparse, urlunparse
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

# Browser-like headers; some OneDrive endpoints reject anonymous share calls without a UA.
_CONSUMER_ONEDRIVE_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "application/json",
}


# ---------------------------------------------------------------------------
# Google Drive
# ---------------------------------------------------------------------------

def list_and_download_gdrive(credentials_json: dict, folder_id: str, progress_callback=None):
    """
    Download all PDFs from a Google Drive folder (recursively).

    Args:
        credentials_json: Service account JSON key (parsed dict)
        folder_id: Google Drive folder ID
        progress_callback: fn(current, total, filename)

    Returns:
        dict of {filename: bytes}
    """
    creds = service_account.Credentials.from_service_account_info(
        credentials_json,
        scopes=['https://www.googleapis.com/auth/drive.readonly'],
    )
    service = build('drive', 'v3', credentials=creds, cache_discovery=False)

    # Collect all PDF file IDs recursively
    pdf_files = []
    _list_pdfs_recursive(service, folder_id, pdf_files)

    if not pdf_files:
        return {}

    total = len(pdf_files)
    result = {}

    for i, (file_id, filename) in enumerate(pdf_files):
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        result[filename] = buf.getvalue()

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return result


def list_and_download_gdrive_public(folder_link: str, progress_callback=None):
    """
    Download all PDFs from a public Google Drive folder link (no auth needed).

    Args:
        folder_link: Google Drive folder URL (must be shared as 'Anyone with the link')
        progress_callback: fn(current, total, filename)

    Returns:
        dict of {filename: bytes}
    """
    import re

    # Extract folder ID from various Google Drive URL formats
    match = re.search(r'/folders/([a-zA-Z0-9_-]+)', folder_link)
    if not match:
        match = re.search(r'id=([a-zA-Z0-9_-]+)', folder_link)
    if not match:
        raise Exception("Could not extract folder ID from the link. Please check the URL.")

    folder_id = match.group(1)

    # Use Google Drive API with API key (public access, no auth)
    # We'll use the files.list endpoint which works for public folders
    api_key_url = "https://www.googleapis.com/drive/v3/files"

    pdf_files = []
    _list_public_pdfs_recursive(api_key_url, folder_id, pdf_files)

    if not pdf_files:
        return {}

    total = len(pdf_files)
    result = {}

    for i, (file_id, filename) in enumerate(pdf_files):
        # Download file content via public link
        download_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media&key="
        # Try direct download approach for public files
        export_url = f"https://drive.google.com/uc?export=download&id={file_id}"
        resp = requests.get(export_url, allow_redirects=True)
        if resp.status_code == 200 and len(resp.content) > 100:
            result[filename] = resp.content
        if progress_callback:
            progress_callback(i + 1, total, filename)

    return result


def _list_public_pdfs_recursive(api_url, folder_id, pdf_files):
    """List PDFs from a public Google Drive folder without auth."""
    page_token = None
    while True:
        params = {
            "q": f"'{folder_id}' in parents and trashed=false",
            "fields": "nextPageToken, files(id, name, mimeType)",
            "pageSize": 1000,
        }
        if page_token:
            params["pageToken"] = page_token

        resp = requests.get(api_url, params=params)

        if resp.status_code == 403 or resp.status_code == 401:
            raise Exception(
                "Cannot access this folder. Make sure it's shared as "
                "'Anyone with the link can view' in Google Drive sharing settings."
            )
        if resp.status_code != 200:
            raise Exception(f"Google Drive API error: {resp.status_code} - {resp.text[:200]}")

        data = resp.json()
        for f in data.get('files', []):
            if f.get('mimeType') == 'application/vnd.google-apps.folder':
                _list_public_pdfs_recursive(api_url, f['id'], pdf_files)
            elif f.get('name', '').lower().endswith('.pdf'):
                pdf_files.append((f['id'], f['name']))

        page_token = data.get('nextPageToken')
        if not page_token:
            break


def _list_pdfs_recursive(service, folder_id, pdf_files):
    """Recursively list all PDFs in a Google Drive folder."""
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name, mimeType)",
            pageSize=1000,
            pageToken=page_token,
        ).execute()

        for f in resp.get('files', []):
            if f['mimeType'] == 'application/vnd.google-apps.folder':
                _list_pdfs_recursive(service, f['id'], pdf_files)
            elif f['name'].lower().endswith('.pdf'):
                pdf_files.append((f['id'], f['name']))

        page_token = resp.get('nextPageToken')
        if not page_token:
            break


# ---------------------------------------------------------------------------
# OneDrive / SharePoint (Microsoft Graph API)
# ---------------------------------------------------------------------------

def list_and_download_onedrive(client_id: str, client_secret: str, tenant_id: str,
                                folder_path: str, progress_callback=None):
    """
    Download all PDFs from a OneDrive/SharePoint folder.

    Args:
        client_id: Azure app registration client ID
        client_secret: Azure app client secret
        tenant_id: Azure tenant ID
        folder_path: Path within the drive, e.g. "Invoices" or "Documents/Invoices"
        progress_callback: fn(current, total, filename)

    Returns:
        dict of {filename: bytes}
    """
    try:
        import msal
    except ImportError as e:
        raise ImportError(
            "Install the msal package for authenticated OneDrive/Graph access: pip install msal"
        ) from e

    # Authenticate
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )
    token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in token_response:
        error = token_response.get("error_description", "Authentication failed")
        raise Exception(f"OneDrive auth failed: {error}")

    headers = {"Authorization": f"Bearer {token_response['access_token']}"}

    # List all PDFs in the folder (recursively)
    pdf_files = []
    _list_onedrive_pdfs(headers, folder_path, pdf_files)

    if not pdf_files:
        return {}

    total = len(pdf_files)
    result = {}

    for i, (download_url, filename) in enumerate(pdf_files):
        resp = requests.get(download_url, headers=headers)
        resp.raise_for_status()
        result[filename] = resp.content

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return result


def _list_onedrive_pdfs(headers, folder_path, pdf_files):
    """Recursively list all PDFs in a OneDrive folder via Graph API."""
    folder_path = folder_path.strip('/')
    # Use /me/drive for personal OneDrive, but with app-only auth we need /users/{id}/drive
    # For simplicity, use the /drives endpoint or let user specify drive ID
    # The most common pattern: list children of a folder by path
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{folder_path}:/children"

    while url:
        resp = requests.get(url, headers=headers, params={"$top": "999"})
        if resp.status_code != 200:
            # Try without /me (app-only permissions)
            # Fall back to searching all drives
            raise Exception(
                f"Cannot access folder '{folder_path}'. "
                f"Status {resp.status_code}: {resp.json().get('error', {}).get('message', resp.text)}"
            )

        data = resp.json()
        for item in data.get('value', []):
            if 'folder' in item:
                # Recurse into subfolders
                child_path = f"{folder_path}/{item['name']}"
                _list_onedrive_pdfs(headers, child_path, pdf_files)
            elif item.get('name', '').lower().endswith('.pdf'):
                download_url = item.get('@microsoft.graph.downloadUrl')
                if download_url:
                    pdf_files.append((download_url, item['name']))

        url = data.get('@odata.nextLink')


def _normalize_onedrive_share_url(share_url: str) -> str:
    """
    Resolve consumer OneDrive wrapper links to the real share URL.

    Copying from the browser often yields onedrive.live.com/?redeem=<base64>
    where the payload decodes to https://1drv.ms/... . The OneDrive sharing
    API token must be built from that inner URL, not the wrapper.
    """
    share_url = share_url.strip()
    parsed = urlparse(share_url)
    if "onedrive.live.com" not in (parsed.netloc or "").lower():
        return share_url

    qs = parse_qs(parsed.query)
    redeem = (qs.get("redeem") or [None])[0]
    if not redeem:
        return share_url

    try:
        padded = redeem + "=" * (-len(redeem) % 4)
        inner = base64.b64decode(padded).decode("utf-8").strip()
        if inner.startswith("http"):
            return inner
    except Exception:
        pass

    return share_url


def _raise_if_onedrive_address_bar_url(share_url: str) -> None:
    """
    onedrive.live.com/?id=...&cid=... is the signed-in web app URL, not an
    anonymous sharing link. Encoding it produces UnauthenticatedVroomException.
    """
    parsed = urlparse(share_url.strip())
    if "onedrive.live.com" not in (parsed.netloc or "").lower():
        return
    qs = parse_qs(parsed.query)
    if qs.get("redeem"):
        return
    if qs.get("id"):
        raise ValueError(
            "That OneDrive URL is from the browser address bar (?id=…), not a sharing link. "
            "The app cannot open it without your Microsoft login.\n\n"
            "Use this instead: open the folder in OneDrive → Share → "
            "“Anyone with the link can view” → Copy link → paste the URL that starts with "
            "https://1drv.ms/ (or paste the long onedrive.live.com link that includes redeem=…)."
        )


def _slim_onedrive_redir_url(url: str) -> str:
    """
    Strip bulky query params (redeem=...) from onedrive.live.com/redir URLs so
    the base64-encoded share token does not exceed maxUrlLength.
    """
    p = urlparse(url)
    q = parse_qs(p.query, keep_blank_values=True)
    keep = {}
    for k in ("resid", "cid", "ithint", "authkey", "e"):
        if k in q and q[k]:
            keep[k] = q[k][0]
    return urlunparse(p._replace(query=urlencode(keep)))


def _follow_short_redirect(share_url: str) -> tuple[str | None, bool]:
    """
    Return (redir_url_or_None, migrated_to_spo).

    Resolves 1drv.ms short links once; detects the migratedtospo=true flag that
    indicates the OneDrive account has been moved to SharePoint Online and is
    unreachable via the anonymous consumer share API.
    """
    try:
        r = requests.head(
            share_url,
            allow_redirects=False,
            headers=_CONSUMER_ONEDRIVE_HEADERS,
            timeout=30,
        )
    except Exception:
        return None, False

    if r.status_code not in (301, 302, 303, 307, 308):
        return None, False

    loc = r.headers.get("Location") or r.headers.get("location")
    if not loc:
        return None, False

    expanded = urljoin(share_url, loc)
    p = urlparse(expanded)
    host = (p.netloc or "").lower()
    if "onedrive.live.com" not in host:
        return expanded, False

    q = parse_qs(p.query)
    migrated = (q.get("migratedtospo") or [""])[0].lower() == "true"
    return expanded, migrated


def _onedrive_share_encoding_candidates(share_url: str) -> list[str]:
    """
    Build URL strings to feed into the u!{base64url} share token.
    """
    share_url = share_url.strip()
    seen: set[str] = set()
    out: list[str] = []

    def add(u: str) -> None:
        u = u.strip()
        if u and u not in seen:
            seen.add(u)
            out.append(u)

    add(share_url)

    host = (urlparse(share_url).netloc or "").lower()
    if "1drv.ms" in host or "1drv.st" in host:
        expanded, _ = _follow_short_redirect(share_url)
        if expanded and "onedrive.live.com" in (urlparse(expanded).netloc or "").lower():
            add(_slim_onedrive_redir_url(expanded))
            add(expanded)

    return out


class OneDriveMigratedToSpoError(Exception):
    """The shared folder has been migrated to SharePoint Online and the
    anonymous consumer share API cannot reach it."""


def list_and_download_onedrive_link(share_url: str, progress_callback=None):
    """
    Download all PDFs from a OneDrive sharing link (no auth needed for public links).

    Args:
        share_url: OneDrive/SharePoint sharing URL (1drv.ms, onedrive.live.com/?redeem=..., etc.)
        progress_callback: fn(current, total, filename)

    Returns:
        dict of {filename: bytes}
    """
    share_url = _normalize_onedrive_share_url(share_url)
    _raise_if_onedrive_address_bar_url(share_url)

    host = (urlparse(share_url).netloc or "").lower()
    if "1drv.ms" in host or "1drv.st" in host:
        _, migrated = _follow_short_redirect(share_url)
        if migrated:
            raise OneDriveMigratedToSpoError(
                "This OneDrive account has been migrated to SharePoint Online "
                "(migratedtospo=true). Microsoft no longer allows anonymous, "
                "public-link downloads for these folders — even when the link "
                "is shared as “Anyone with the link can view.”\n\n"
                "Please use one of these instead:\n"
                "• Upload Files / ZIP in this app (always works), or\n"
                "• Put the PDFs in a Google Drive folder shared as "
                "“Anyone with the link can view” and use the Google Drive tab, or\n"
                "• Ask a developer to add OneDrive OAuth sign-in to the app."
            )

    last_error: Exception | None = None
    pdf_files: list[tuple[str, str]] = []

    for candidate in _onedrive_share_encoding_candidates(share_url):
        encoded = base64.urlsafe_b64encode(candidate.encode()).decode().rstrip('=')
        token = 'u!' + encoded
        api_base = f"https://api.onedrive.com/v1.0/shares/{token}"
        try:
            batch: list[tuple[str, str]] = []
            _list_share_pdfs(api_base, "/root/children", batch)
            pdf_files = batch
            last_error = None
            break
        except Exception as e:
            last_error = e
            continue

    if last_error is not None:
        raise last_error

    if not pdf_files:
        return {}

    total = len(pdf_files)
    result = {}

    for i, (download_url, filename) in enumerate(pdf_files):
        resp = requests.get(
            download_url,
            headers=_CONSUMER_ONEDRIVE_HEADERS,
            timeout=120,
        )
        resp.raise_for_status()
        result[filename] = resp.content

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return result


def _list_share_pdfs(api_base, path, pdf_files):
    """List PDFs from a OneDrive sharing link."""
    url = api_base + path
    resp = requests.get(
        url,
        params={"$top": "999"},
        headers=_CONSUMER_ONEDRIVE_HEADERS,
        timeout=60,
    )

    if resp.status_code != 200:
        err_body = resp.text
        try:
            err_body = resp.json().get("error", {}).get("message", err_body)
        except Exception:
            pass
        hint = ""
        err_s = str(err_body)
        if "UnauthenticatedVroomException" in err_s or err_s.strip() == "Unauthenticated":
            hint = (
                " Microsoft often returns this if the link is not really “Anyone with the link”, "
                "the link expired, or the folder was moved. "
                "Try creating a fresh Share link, or use Upload Files / a public Google Drive folder instead."
            )
        if "UnauthenticatedVroomException" in err_s:
            hint += (
                " If you used onedrive.live.com/?id=… from the address bar, use Share → Copy link "
                "(https://1drv.ms/… or redeem=…) instead."
            )
        raise Exception(
            f"Cannot access shared folder. Make sure the link is set to "
            f"'Anyone with the link can view'. Error: {err_body}{hint}"
        )

    data = resp.json()
    for item in data.get('value', []):
        if 'folder' in item:
            child_path = f"/root/children/{item['id']}/children" if '/children' in path else f"/{item['id']}/children"
            # Use item ID to get children
            child_url_path = f"/items/{item['id']}/children"
            _list_share_pdfs(api_base, child_url_path, pdf_files)
        elif item.get('name', '').lower().endswith('.pdf'):
            download_url = item.get('@content.downloadUrl') or item.get('@microsoft.graph.downloadUrl')
            if download_url:
                pdf_files.append((download_url, item['name']))

    next_link = data.get('@odata.nextLink')
    if next_link:
        # For pagination, use the full URL
        resp2 = requests.get(next_link, headers=_CONSUMER_ONEDRIVE_HEADERS, timeout=60)
        if resp2.status_code == 200:
            data2 = resp2.json()
            for item in data2.get('value', []):
                if item.get('name', '').lower().endswith('.pdf'):
                    download_url = item.get('@content.downloadUrl') or item.get('@microsoft.graph.downloadUrl')
                    if download_url:
                        pdf_files.append((download_url, item['name']))
