"""
Microsoft OAuth (authorization code + PKCE) and Microsoft Graph downloads
for OneDrive / SharePoint-backed personal folders (including migratedtospo).
"""

from __future__ import annotations

import base64
import os
import time
from typing import Any, Callable

import msal
import requests

GRAPH = "https://graph.microsoft.com/v1.0"

# Delegated permissions — user signs in; works for shares the user can open.
GRAPH_SCOPES = ["Files.Read.All", "User.Read", "offline_access"]


def build_msal_app(cfg: dict) -> msal.ClientApplication:
    tenant = cfg.get("tenant_id") or "common"
    authority = f"https://login.microsoftonline.com/{tenant}"
    client_id = cfg["client_id"]
    secret = cfg.get("client_secret")
    if secret:
        return msal.ConfidentialClientApplication(
            client_id,
            client_credential=secret,
            authority=authority,
        )
    return msal.PublicClientApplication(client_id, authority=authority)


def start_auth_code_flow(app: msal.ClientApplication, redirect_uri: str) -> dict:
    return app.initiate_auth_code_flow(
        GRAPH_SCOPES,
        redirect_uri=redirect_uri,
    )


def acquire_token_from_auth_response(
    app: msal.ClientApplication,
    flow: dict,
    auth_response: dict[str, str],
) -> dict[str, Any]:
    return app.acquire_token_by_auth_code_flow(
        flow,
        auth_response,
        scopes=GRAPH_SCOPES,
    )


def refresh_tokens(app: msal.ClientApplication, refresh_token: str) -> dict[str, Any]:
    return app.acquire_token_by_refresh_token(refresh_token, GRAPH_SCOPES)


def get_valid_access_token(
    cfg: dict,
    auth: dict[str, Any] | None,
    expires_at: float | None,
) -> tuple[str | None, dict[str, Any] | None, float | None]:
    """Return (access_token, updated_auth_dict, new_expires_at) or Nones if re-login needed."""
    if not auth or "access_token" not in auth:
        return None, None, None

    now = time.time()
    if expires_at and now < expires_at - 120:
        return auth["access_token"], auth, expires_at

    rt = auth.get("refresh_token")
    if not rt:
        return auth["access_token"], auth, expires_at

    app = build_msal_app(cfg)
    result = refresh_tokens(app, rt)
    if "access_token" in result:
        merged = {**auth, **result}
        new_expires = now + int(result.get("expires_in", 3600))
        return merged["access_token"], merged, new_expires
    return None, None, None


def _encode_graph_share_id(sharing_url: str) -> str:
    enc = base64.urlsafe_b64encode(sharing_url.encode()).decode().rstrip("=")
    return "u!" + enc


def _unique_basename(filename: str, seen: set[str]) -> str:
    if filename not in seen:
        seen.add(filename)
        return filename
    base, ext = os.path.splitext(filename)
    n = 2
    while True:
        cand = f"{base}_{n}{ext}"
        if cand not in seen:
            seen.add(cand)
            return cand
        n += 1


def _drive_context_from_item(item: dict) -> tuple[str, str]:
    """Resolve drive_id and item_id for /drives/.../items/... calls."""
    if "remoteItem" in item:
        ri = item["remoteItem"]
        pr = ri.get("parentReference") or {}
        drive_id = pr.get("driveId")
        item_id = ri.get("id")
        if drive_id and item_id:
            return drive_id, item_id
    pr = item.get("parentReference") or {}
    drive_id = pr.get("driveId")
    item_id = item.get("id")
    if not drive_id or not item_id:
        raise ValueError("Could not resolve drive from shared link response.")
    return drive_id, item_id


def _graph_get_json(url: str, headers: dict, timeout: int = 60) -> dict:
    r = requests.get(url, headers=headers, timeout=timeout)
    if r.status_code != 200:
        try:
            err = r.json().get("error", {})
            msg = err.get("message", r.text[:300])
        except Exception:
            msg = r.text[:300]
        raise RuntimeError(f"Microsoft Graph error ({r.status_code}): {msg}")
    return r.json()


def _list_children_paged(url: str, headers: dict) -> list[dict]:
    items: list[dict] = []
    while url:
        data = _graph_get_json(url, headers)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def _collect_pdf_nodes(
    drive_id: str,
    item_id: str,
    headers: dict,
    rel_path: str,
    out: list[tuple[str, str, str]],
) -> None:
    base = f"{GRAPH}/drives/{drive_id}/items/{item_id}/children"
    for it in _list_children_paged(base, headers):
        name = it.get("name") or ""
        if it.get("folder"):
            sub = f"{rel_path}/{name}" if rel_path else name
            _collect_pdf_nodes(drive_id, it["id"], headers, sub, out)
            continue
        if not name.lower().endswith(".pdf"):
            continue
        dl = it.get("@microsoft.graph.downloadUrl")
        if not dl:
            detail = _graph_get_json(
                f"{GRAPH}/drives/{drive_id}/items/{it['id']}",
                headers,
            )
            dl = detail.get("@microsoft.graph.downloadUrl")
        if not dl:
            continue
        display_name = f"{rel_path}/{name}" if rel_path else name
        out.append((display_name, dl, display_name))


def download_shared_folder_via_graph(
    access_token: str,
    share_url: str,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> dict[str, bytes]:
    """
    Download all PDFs from a shared OneDrive / SharePoint folder using Graph + user token.

    share_url: 1drv.ms, onedrive.live.com/?redeem=..., or redir URL after normalization.
    """
    from cloud_sources import (  # local import avoids circular deps at module load
        _normalize_onedrive_share_url,
        _raise_if_onedrive_address_bar_url,
    )

    raw = share_url.strip()
    raw = _normalize_onedrive_share_url(raw)
    _raise_if_onedrive_address_bar_url(raw)

    share_id = _encode_graph_share_id(raw)
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
    }

    item = _graph_get_json(f"{GRAPH}/shares/{share_id}/driveItem", headers)

    if item.get("folder"):
        drive_id, root_id = _drive_context_from_item(item)
        nodes: list[tuple[str, str, str]] = []
        _collect_pdf_nodes(drive_id, root_id, headers, "", nodes)
    else:
        name = item.get("name") or "file.pdf"
        if not name.lower().endswith(".pdf"):
            raise ValueError(
                "The shared link points to a file that is not a PDF, and is not a folder."
            )
        dl = item.get("@microsoft.graph.downloadUrl")
        if not dl:
            drive_id, fid = _drive_context_from_item(item)
            detail = _graph_get_json(
                f"{GRAPH}/drives/{drive_id}/items/{fid}",
                headers,
            )
            dl = detail.get("@microsoft.graph.downloadUrl")
        if not dl:
            raise RuntimeError("Microsoft Graph did not return a download URL for this item.")
        nodes = [(name, dl, name)]

    if not nodes:
        return {}

    total = len(nodes)
    result: dict[str, bytes] = {}
    seen_names: set[str] = set()

    for i, (_display, dl_url, logical_name) in enumerate(nodes):
        if progress_callback:
            progress_callback(i, total, logical_name)
        r = requests.get(dl_url, timeout=120)
        r.raise_for_status()
        key = _unique_basename(os.path.basename(logical_name.replace("\\", "/")), seen_names)
        result[key] = r.content

    if progress_callback:
        progress_callback(total, total, "done")

    return result
