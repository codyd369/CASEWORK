"""
Microsoft Graph API client for SharePoint file operations.

Handles authentication (both daemon and delegated flows), file download/upload,
folder listing, and link generation for opening documents in native apps.
"""

import io
import json
import logging
import os
import time
from pathlib import Path
from typing import Optional

import msal
import requests

from config.settings import (
    AZURE_AUTHORITY,
    AZURE_CLIENT_ID,
    AZURE_CLIENT_SECRET,
    AZURE_SCOPES,
    AZURE_SCOPES_DELEGATED,
    AZURE_TENANT_ID,
    BASE_FOLDER_PATH,
    SHAREPOINT_SITE_URL,
)

logger = logging.getLogger(__name__)

# Token cache file for interactive auth
TOKEN_CACHE_FILE = os.path.join(os.path.expanduser("~"), ".casework_token_cache.json")


class GraphAPIClient:
    """Authenticated Microsoft Graph API client for SharePoint operations."""

    GRAPH_BASE = "https://graph.microsoft.com/v1.0"

    def __init__(self, use_delegated_auth: bool = False):
        """
        Initialize the Graph API client.

        Args:
            use_delegated_auth: If True, use interactive browser-based auth
                (required for the companion app so facts are attributed to the user).
                If False, use client-credentials (daemon) flow for background scripts.
        """
        self.use_delegated_auth = use_delegated_auth
        self._token: Optional[str] = None
        self._token_expiry: float = 0
        self._site_id: Optional[str] = None
        self._drive_id: Optional[str] = None
        self._user_name: Optional[str] = None

        # Build the MSAL app
        if use_delegated_auth:
            self._cache = msal.SerializableTokenCache()
            if os.path.exists(TOKEN_CACHE_FILE):
                self._cache.deserialize(open(TOKEN_CACHE_FILE).read())
            self._app = msal.PublicClientApplication(
                AZURE_CLIENT_ID,
                authority=AZURE_AUTHORITY,
                token_cache=self._cache,
            )
        else:
            self._app = msal.ConfidentialClientApplication(
                AZURE_CLIENT_ID,
                authority=AZURE_AUTHORITY,
                client_credential=AZURE_CLIENT_SECRET,
            )

    # ── Authentication ──

    def _get_token(self) -> str:
        """Get a valid access token, refreshing if needed."""
        if self._token and time.time() < self._token_expiry - 60:
            return self._token

        if self.use_delegated_auth:
            result = self._acquire_delegated_token()
        else:
            result = self._app.acquire_token_for_client(scopes=AZURE_SCOPES)

        if "access_token" in result:
            self._token = result["access_token"]
            self._token_expiry = time.time() + result.get("expires_in", 3600)
            # Persist token cache for delegated auth
            if self.use_delegated_auth and self._cache.has_state_changed:
                with open(TOKEN_CACHE_FILE, "w") as f:
                    f.write(self._cache.serialize())
            return self._token
        else:
            error = result.get("error_description", result.get("error", "Unknown error"))
            raise RuntimeError(f"Failed to acquire token: {error}")

    def _acquire_delegated_token(self) -> dict:
        """Acquire token via interactive browser flow (with cache)."""
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(
                AZURE_SCOPES_DELEGATED, account=accounts[0]
            )
            if result and "access_token" in result:
                return result

        # Fall back to interactive login
        return self._app.acquire_token_interactive(scopes=AZURE_SCOPES_DELEGATED)

    def _headers(self) -> dict:
        """Return authorization headers."""
        return {"Authorization": f"Bearer {self._get_token()}"}

    # ── Site & Drive Resolution ──

    def get_site_id(self) -> str:
        """Resolve the SharePoint site ID from the site URL."""
        if self._site_id:
            return self._site_id

        # Parse the site URL to extract hostname and site path
        from urllib.parse import urlparse
        parsed = urlparse(SHAREPOINT_SITE_URL)
        hostname = parsed.hostname  # e.g. blueprintconstructionlaw.sharepoint.com
        site_path = parsed.path     # e.g. /sites/PCLFormosa813

        url = f"{self.GRAPH_BASE}/sites/{hostname}:{site_path}"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        self._site_id = resp.json()["id"]
        logger.info(f"Resolved site ID: {self._site_id}")
        return self._site_id

    def get_drive_id(self) -> str:
        """Get the default document library drive ID for the site."""
        if self._drive_id:
            return self._drive_id

        site_id = self.get_site_id()
        url = f"{self.GRAPH_BASE}/sites/{site_id}/drives"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        drives = resp.json().get("value", [])

        # Find the "Documents" drive (default Teams document library)
        for drive in drives:
            if drive.get("name") == "Documents" or "Shared Documents" in drive.get("webUrl", ""):
                self._drive_id = drive["id"]
                break
        if not self._drive_id and drives:
            self._drive_id = drives[0]["id"]

        if not self._drive_id:
            raise RuntimeError("Could not find a document library drive.")

        logger.info(f"Resolved drive ID: {self._drive_id}")
        return self._drive_id

    def get_current_user_name(self) -> str:
        """Get the display name of the currently authenticated user (delegated auth only)."""
        if self._user_name:
            return self._user_name
        if not self.use_delegated_auth:
            return "System"

        url = f"{self.GRAPH_BASE}/me"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        self._user_name = resp.json().get("displayName", "Unknown User")
        return self._user_name

    # ── File Path Helpers ──

    def _item_path(self, relative_path: str) -> str:
        """Build the Graph API path for a file/folder within the base folder."""
        drive_id = self.get_drive_id()
        full_path = f"{BASE_FOLDER_PATH}/{relative_path}".strip("/")
        return f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{full_path}"

    # ── File Operations ──

    def download_file(self, relative_path: str) -> bytes:
        """Download a file from SharePoint and return its content as bytes."""
        url = f"{self._item_path(relative_path)}:/content"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.content

    def download_file_to_stream(self, relative_path: str) -> io.BytesIO:
        """Download a file and return as a BytesIO stream (for openpyxl)."""
        return io.BytesIO(self.download_file(relative_path))

    def upload_file(self, relative_path: str, content: bytes) -> dict:
        """
        Upload (create or overwrite) a file in SharePoint.

        For files up to 4MB, uses simple upload. For larger files,
        uses an upload session (resumable upload).
        """
        if len(content) <= 4 * 1024 * 1024:
            return self._simple_upload(relative_path, content)
        else:
            return self._resumable_upload(relative_path, content)

    def _simple_upload(self, relative_path: str, content: bytes) -> dict:
        """Simple upload for files ≤ 4MB."""
        url = f"{self._item_path(relative_path)}:/content"
        headers = self._headers()
        headers["Content-Type"] = "application/octet-stream"
        resp = requests.put(url, headers=headers, data=content)
        resp.raise_for_status()
        return resp.json()

    def _resumable_upload(self, relative_path: str, content: bytes) -> dict:
        """Resumable upload session for files > 4MB."""
        # Create upload session
        url = f"{self._item_path(relative_path)}:/createUploadSession"
        resp = requests.post(url, headers=self._headers(), json={
            "item": {"@microsoft.graph.conflictBehavior": "replace"}
        })
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]

        # Upload in 5MB chunks
        chunk_size = 5 * 1024 * 1024
        total = len(content)
        for start in range(0, total, chunk_size):
            end = min(start + chunk_size, total)
            chunk = content[start:end]
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {start}-{end - 1}/{total}",
            }
            resp = requests.put(upload_url, headers=headers, data=chunk)
            resp.raise_for_status()

        return resp.json()

    def list_folder(self, relative_path: str) -> list[dict]:
        """
        List items in a SharePoint folder.

        Returns a list of dicts with keys: name, id, webUrl, size, file/folder.
        Handles pagination automatically.
        """
        drive_id = self.get_drive_id()
        full_path = f"{BASE_FOLDER_PATH}/{relative_path}".strip("/")
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root:/{full_path}:/children"
        params = {"$top": 200}

        all_items = []
        while url:
            resp = requests.get(url, headers=self._headers(), params=params)
            resp.raise_for_status()
            data = resp.json()
            all_items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = {}  # nextLink already contains query params

        return all_items

    def list_files_in_folder(self, relative_path: str) -> dict[str, dict]:
        """
        List all files in a folder, keyed by filename (without extension).

        Returns: {stem: {name, webUrl, extension, id}, ...}
        """
        items = self.list_folder(relative_path)
        result = {}
        for item in items:
            if "file" in item:  # Skip subfolders
                name = item["name"]
                stem = Path(name).stem
                result[stem] = {
                    "name": name,
                    "webUrl": item.get("webUrl", ""),
                    "extension": Path(name).suffix.lower(),
                    "id": item["id"],
                }
        return result

    def get_file_open_url(self, relative_path: str) -> str:
        """
        Get a URL that opens the file in the user's native desktop application.

        Uses the SharePoint webUrl — when opened, SharePoint/Office prompts
        to open in the desktop app. For non-Office files (PDFs), the browser
        download triggers the native app.
        """
        url = f"{self._item_path(relative_path)}"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        data = resp.json()
        web_url = data.get("webUrl", "")

        # For Office docs, append ?web=0 to force desktop app opening
        ext = Path(relative_path).suffix.lower()
        if ext in (".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt"):
            return f"{web_url}?web=0"
        return web_url

    def get_file_metadata(self, relative_path: str) -> dict:
        """Get metadata (id, name, size, webUrl, etc.) for a file."""
        url = f"{self._item_path(relative_path)}"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.json()

    def file_exists(self, relative_path: str) -> bool:
        """Check if a file exists in SharePoint."""
        try:
            self.get_file_metadata(relative_path)
            return True
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                return False
            raise

    def search_files(self, query: str) -> list[dict]:
        """Search for files in the drive by name or content."""
        drive_id = self.get_drive_id()
        url = f"{self.GRAPH_BASE}/drives/{drive_id}/root/search(q='{query}')"
        resp = requests.get(url, headers=self._headers())
        resp.raise_for_status()
        return resp.json().get("value", [])
