"""
Google Drive uploader for categorized invoices.
Uploads files in CustomerName/Year/Month/ folder structure.

Uses OAuth 2.0 user delegation (a stored refresh token), not a service account.
Service accounts have no "My Drive" storage quota and cannot upload to a personal
Drive folder; OAuth user creds upload as the signed-in user, who owns the folder.

Required oauth_config keys: client_id, client_secret, refresh_token.
See SETUP_GDRIVE.md for the one-time setup that produces these values.
"""

import io
from typing import Mapping, Any

from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/drive"]
TOKEN_URI = "https://oauth2.googleapis.com/token"


def _build_user_credentials(oauth_config: Mapping[str, Any]) -> Credentials:
    """Build OAuth user credentials from a stored refresh token. Auto-refreshes on use."""
    return Credentials(
        token=None,
        refresh_token=oauth_config["refresh_token"],
        client_id=oauth_config["client_id"],
        client_secret=oauth_config["client_secret"],
        token_uri=TOKEN_URI,
        scopes=SCOPES,
    )


def _get_service(oauth_config: Mapping[str, Any]):
    """Build Drive API service from OAuth user credentials."""
    creds = _build_user_credentials(oauth_config)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _drive_name_escape(s: str) -> str:
    """Escape a file name for use inside Drive API query string quotes."""
    return s.replace("\\", "\\\\").replace("'", "\\'")


def _find_or_create_folder(service, name: str, parent_id: str) -> str:
    """Find an existing folder or create one under parent_id (supports shared drives)."""
    n = _drive_name_escape(name)
    query = (
        f"name = '{n}' and mimeType = 'application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed = false"
    )
    results = (
        service.files()
        .list(
            q=query,
            fields="files(id, name)",
            pageSize=10,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    files = results.get("files", [])
    if files:
        return files[0]["id"]

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = (
        service.files()
        .create(
            body=metadata,
            fields="id",
            supportsAllDrives=True,
        )
        .execute()
    )
    return folder["id"]


def _find_existing_file_id(service, filename: str, parent_id: str) -> str | None:
    """Return file id if a non-trashed file with the same name exists in parent."""
    n = _drive_name_escape(filename)
    query = f"name = '{n}' and '{parent_id}' in parents and trashed = false"
    results = (
        service.files()
        .list(
            q=query,
            fields="files(id, name, mimeType)",
            pageSize=5,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        )
        .execute()
    )
    for f in results.get("files", []):
        if f.get("mimeType") == "application/vnd.google-apps.folder":
            continue
        return f.get("id")
    return None


def upload_to_drive(
    oauth_config: Mapping[str, Any],
    parent_folder_id: str,
    ok_results: list[dict],
    uploaded_files: dict[str, bytes],
    progress_callback=None,
    skip_if_exists: bool = True,
):
    """
    Upload categorized invoices to Google Drive as the OAuth user.

    Args:
        oauth_config: dict with client_id, client_secret, refresh_token
        parent_folder_id: Root folder owned by (or shared with) the OAuth user
        ok_results: List of result dicts with status='ok'
        uploaded_files: {filename: pdf_bytes}
        progress_callback: fn(current, total, filename) after each *attempt*
        skip_if_exists: If True, skip when file already in destination

    Returns:
        dict with uploaded, skipped_existing, total_attempts, parent_folder_id
    """
    if not ok_results:
        return {
            "uploaded": 0,
            "skipped_existing": 0,
            "skipped_names": [],
            "root_folder_id": parent_folder_id,
            "parent_folder_id": parent_folder_id,
        }

    service = _get_service(oauth_config)

    root_id = parent_folder_id

    folder_cache: dict[str, str] = {}
    uploaded_count = 0
    skipped_existing = 0
    skipped_names: list[str] = []
    total = len(ok_results)

    for i, r in enumerate(ok_results):
        filename = r["filename"]
        customer = r["customer_clean"]
        pdf_bytes = uploaded_files.get(filename)
        if not pdf_bytes:
            continue

        if r.get("year") and r.get("month"):
            path_parts = [customer, str(r["year"]), f"{r['month']:02d}"]
        else:
            path_parts = [customer, "_UNKNOWN_DATE"]

        cache_key = "/".join(path_parts)
        if cache_key in folder_cache:
            target_folder_id = folder_cache[cache_key]
        else:
            current_parent = root_id
            for j, part in enumerate(path_parts):
                partial_key = "/".join(path_parts[: j + 1])
                if partial_key in folder_cache:
                    current_parent = folder_cache[partial_key]
                else:
                    current_parent = _find_or_create_folder(service, part, current_parent)
                    folder_cache[partial_key] = current_parent
            target_folder_id = current_parent
            folder_cache[cache_key] = target_folder_id

        if skip_if_exists and _find_existing_file_id(service, filename, target_folder_id):
            skipped_existing += 1
            skipped_names.append(filename)
            if progress_callback:
                progress_callback(i + 1, total, filename)
            continue

        file_metadata = {
            "name": filename,
            "parents": [target_folder_id],
        }
        media = MediaIoBaseUpload(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            resumable=True,
        )
        service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id",
            supportsAllDrives=True,
        ).execute()

        uploaded_count += 1
        if progress_callback:
            progress_callback(i + 1, total, filename)

    return {
        "uploaded": uploaded_count,
        "skipped_existing": skipped_existing,
        "skipped_names": skipped_names,
        "root_folder_id": root_id,
        "parent_folder_id": root_id,
    }
