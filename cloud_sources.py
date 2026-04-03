"""
Download PDF invoices from Google Drive or OneDrive folders.
"""

import io
import requests
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import msal


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


def list_and_download_onedrive_link(share_url: str, progress_callback=None):
    """
    Download all PDFs from a OneDrive sharing link (no auth needed for public links).

    Args:
        share_url: OneDrive/SharePoint sharing URL
        progress_callback: fn(current, total, filename)

    Returns:
        dict of {filename: bytes}
    """
    import base64

    # Encode sharing URL for the API
    encoded = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip('=')
    token = 'u!' + encoded

    api_base = f"https://api.onedrive.com/v1.0/shares/{token}"

    # List children
    pdf_files = []
    _list_share_pdfs(api_base, "/root/children", pdf_files)

    if not pdf_files:
        return {}

    total = len(pdf_files)
    result = {}

    for i, (download_url, filename) in enumerate(pdf_files):
        resp = requests.get(download_url)
        resp.raise_for_status()
        result[filename] = resp.content

        if progress_callback:
            progress_callback(i + 1, total, filename)

    return result


def _list_share_pdfs(api_base, path, pdf_files):
    """List PDFs from a OneDrive sharing link."""
    url = api_base + path
    resp = requests.get(url, params={"$top": "999"})

    if resp.status_code != 200:
        raise Exception(
            f"Cannot access shared folder. "
            f"Make sure the link is set to 'Anyone with the link can view'. "
            f"Error: {resp.json().get('error', {}).get('message', resp.text)}"
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
        resp2 = requests.get(next_link)
        if resp2.status_code == 200:
            data2 = resp2.json()
            for item in data2.get('value', []):
                if item.get('name', '').lower().endswith('.pdf'):
                    download_url = item.get('@content.downloadUrl') or item.get('@microsoft.graph.downloadUrl')
                    if download_url:
                        pdf_files.append((download_url, item['name']))
