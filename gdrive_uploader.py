"""
Google Drive uploader for categorized invoices.
Uploads files in CustomerName/Year/Month/ folder structure.
"""

import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive.file']


def _get_service(credentials_json: dict):
    """Build Drive API service from service account credentials."""
    creds = service_account.Credentials.from_service_account_info(
        credentials_json, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def _find_or_create_folder(service, name: str, parent_id: str) -> str:
    """Find an existing folder or create one under parent_id."""
    query = (
        f"name='{name}' and mimeType='application/vnd.google-apps.folder' "
        f"and '{parent_id}' in parents and trashed=false"
    )
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']

    metadata = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id],
    }
    folder = service.files().create(body=metadata, fields='id').execute()
    return folder['id']


def upload_to_drive(
    credentials_json: dict,
    parent_folder_id: str,
    ok_results: list[dict],
    uploaded_files: dict[str, bytes],
    progress_callback=None,
):
    """
    Upload categorized invoices to Google Drive.

    Args:
        credentials_json: Service account JSON key (parsed dict)
        parent_folder_id: Google Drive folder ID to upload into
        ok_results: List of result dicts with status='ok'
        uploaded_files: {filename: pdf_bytes}
        progress_callback: fn(current, total, filename) called after each file

    Returns:
        dict with 'uploaded' count and 'root_folder_id'
    """
    service = _get_service(credentials_json)

    # Create root folder "Categorized Invoices"
    root_id = _find_or_create_folder(service, "Categorized Invoices", parent_folder_id)

    # Cache folder IDs to avoid repeated API calls
    folder_cache = {}
    uploaded_count = 0
    total = len(ok_results)

    for i, r in enumerate(ok_results):
        filename = r['filename']
        customer = r['customer_clean']
        pdf_bytes = uploaded_files.get(filename)
        if not pdf_bytes:
            continue

        # Build folder path
        if r['year'] and r['month']:
            path_parts = [customer, str(r['year']), f"{r['month']:02d}"]
        else:
            path_parts = [customer, "_UNKNOWN_DATE"]

        # Create nested folders
        cache_key = "/".join(path_parts)
        if cache_key in folder_cache:
            target_folder_id = folder_cache[cache_key]
        else:
            current_parent = root_id
            for j, part in enumerate(path_parts):
                partial_key = "/".join(path_parts[:j + 1])
                if partial_key in folder_cache:
                    current_parent = folder_cache[partial_key]
                else:
                    current_parent = _find_or_create_folder(service, part, current_parent)
                    folder_cache[partial_key] = current_parent
            target_folder_id = current_parent
            folder_cache[cache_key] = target_folder_id

        # Upload file
        file_metadata = {
            'name': filename,
            'parents': [target_folder_id],
        }
        media = MediaIoBaseUpload(
            io.BytesIO(pdf_bytes),
            mimetype='application/pdf',
            resumable=True,
        )
        service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
        ).execute()

        uploaded_count += 1
        if progress_callback:
            progress_callback(i + 1, total, filename)

    return {
        'uploaded': uploaded_count,
        'root_folder_id': root_id,
    }
