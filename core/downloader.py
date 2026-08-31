import re
import logging
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import requests
from fastapi import HTTPException
from typing import Optional

from core.logger import log_job_message

log = logging.getLogger("ei_stream_server.downloader")

def _handle_drive_confirm_form(session: requests.Session, html_text: str, headers: dict = None, job_id: Optional[str] = None) -> Optional[requests.Response]:
    """Parse Google Drive HTML warning page form (<form id='download-form'>) or download link and submit it."""
    # 1. Try form matching
    action_match = re.search(r'<form[^>]*action="([^"]+)"', html_text)
    if action_match:
        action_url = action_match.group(1).replace("&amp;", "&")
        if action_url.startswith("//"):
            action_url = "https:" + action_url
        elif action_url.startswith("/"):
            action_url = "https://drive.google.com" + action_url
        elif not action_url.startswith("http"):
            action_url = "https://drive.usercontent.google.com/" + action_url

        inputs = re.findall(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"', html_text)
        params = {name: val for name, val in inputs}

        msg = f"Auto-submitting Google Drive confirmation form to: {action_url} (params: {list(params.keys())})"
        log.info(msg)
        if job_id:
            log_job_message(job_id, f"📝 {msg}")
        try:
            resp = session.get(action_url, params=params, headers=headers or {}, stream=True, timeout=180, allow_redirects=True)
            if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", "").lower():
                return resp
        except Exception as e:
            log.warning(f"Error submitting Drive confirmation form: {e}")

    # 2. Try link matching (uc-download-link or confirm link)
    link_match = re.search(r'href="([^"]*(?:confirm|download)[^"]*)"', html_text)
    if link_match:
        link_url = link_match.group(1).replace("&amp;", "&")
        if link_url.startswith("//"):
            link_url = "https:" + link_url
        elif link_url.startswith("/"):
            link_url = "https://drive.google.com" + link_url
        elif not link_url.startswith("http"):
            link_url = "https://drive.google.com/" + link_url

        msg = f"Following Google Drive confirmation link: {link_url[:80]}..."
        log.info(msg)
        if job_id:
            log_job_message(job_id, f"🔗 {msg}")
        try:
            resp = session.get(link_url, headers=headers or {}, stream=True, timeout=180, allow_redirects=True)
            if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", "").lower():
                return resp
        except Exception as e:
            log.warning(f"Error following Drive confirmation link: {e}")

    # 3. Direct usercontent.google.com fallback with confirm=t
    confirm_token_match = re.search(r'confirm=([a-zA-Z0-9_-]+)', html_text) or re.search(r'name="confirm"\s+value="([^"]+)"', html_text)
    confirm_token = confirm_token_match.group(1) if confirm_token_match else "t"
    uuid_match = re.search(r'name="uuid"\s+value="([^"]+)"', html_text)
    uuid_val = uuid_match.group(1) if uuid_match else None
    
    file_id_match = re.search(r'id=([a-zA-Z0-9_-]{25,})', html_text) or re.search(r'name="id"\s+value="([^"]+)"', html_text)
    if file_id_match:
        f_id = file_id_match.group(1)
        fallback_urls = [
            f"https://drive.usercontent.google.com/download?id={f_id}&export=download&confirm={confirm_token}" + (f"&uuid={uuid_val}" if uuid_val else ""),
            f"https://drive.google.com/uc?export=download&confirm={confirm_token}&id={f_id}"
        ]
        for f_url in fallback_urls:
            try:
                log.info(f"Trying usercontent confirmation fallback: {f_url}")
                resp = session.get(f_url, headers=headers or {}, stream=True, timeout=180, allow_redirects=True)
                if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", "").lower():
                    return resp
            except Exception as e:
                log.warning(f"Failed usercontent fallback ({f_url}): {e}")

    return None

def extract_file_id(url_or_id: str) -> str:
    """Extract Google Drive file ID from a full URL, or return as-is if already a bare ID."""
    url_or_id = url_or_id.strip()

    match = re.search(r'/(?:d|files)/([a-zA-Z0-9_-]{25,})', url_or_id)
    if match:
        return match.group(1)

    match = re.search(r'[?&]id=([a-zA-Z0-9_-]{25,})', url_or_id)
    if match:
        return match.group(1)

    if re.match(r'^[a-zA-Z0-9_-]{25,}$', url_or_id):
        return url_or_id

    raise HTTPException(
        status_code=400,
        detail=f"Invalid Google Drive URL or File ID: '{url_or_id[:200]}'"
    )

def extract_token_from_url(url: str) -> Optional[str]:
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    if "access_token" in qs:
        return qs["access_token"][0]
    return None

def is_direct_download_url(url: str) -> bool:
    """Returns True if the URL carries an explicit OAuth access_token or API key or export format."""
    return "access_token" in url or "alt=media" in url or "exportFormat=xlsx" in url or "googleapis.com" in url

def _save_stream_to_file(response: requests.Response, dest_path: Path, job_id: Optional[str] = None) -> Path:
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    total_bytes = 0
    with open(dest_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)
                total_bytes += len(chunk)
    _validate_downloaded_file(dest_path)
    msg = f"Saved downloaded stream to disk: '{dest_path.name}' ({dest_path.stat().st_size / 1024 / 1024:.2f} MB)"
    log.info(msg)
    if job_id:
        log_job_message(job_id, f"💾 {msg}")
    return dest_path

def download_from_url(url: str, dest_path: Path, job_id: Optional[str] = None) -> Path:
    """
    Download a file directly from a fully-formed URL in 32KB stream chunks directly to disk.
    Tries authenticated endpoints if an access_token is present.
    """
    token = extract_token_from_url(url)
    file_id = extract_file_id(url)
    log.info(f"Direct stream download → '{dest_path.name}' (File ID: {file_id}, Auth: {bool(token)})")
    if job_id:
        log_job_message(job_id, f"📥 Initiating stream download for File ID: {file_id}")

    session = requests.Session()
    auth_headers = {"Authorization": f"Bearer {token}"} if token else {}

    # If OAuth token is present, try Google Drive API v3 endpoints first
    if token:
        # Method 1: Google Drive API v3 Files.export (for Google Sheets)
        export_api_url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        try:
            log.info(f"Trying Drive API v3 export for Google Sheet '{file_id}'...")
            resp = session.get(export_api_url, headers=auth_headers, stream=True, timeout=120)
            if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", "").lower():
                return _save_stream_to_file(resp, dest_path, job_id=job_id)
        except Exception as e:
            log.warning(f"Drive API export failed: {e}")

        # Method 2: Google Drive API v3 Files.get alt=media (for binary files)
        media_api_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
        try:
            log.info(f"Trying Drive API v3 alt=media for binary file '{file_id}'...")
            resp = session.get(media_api_url, headers=auth_headers, stream=True, timeout=120)
            if resp.status_code == 200 and "text/html" not in resp.headers.get("Content-Type", "").lower():
                return _save_stream_to_file(resp, dest_path, job_id=job_id)
        except Exception as e:
            log.warning(f"Drive API alt=media failed: {e}")

    # Method 3: Fetch directly from the supplied URL
    try:
        response = session.get(url, headers=auth_headers, stream=True, timeout=120, allow_redirects=True)
        if response.status_code == 200:
            content_type = response.headers.get("Content-Type", "").lower()
            if "text/html" in content_type:
                html_text = response.text
                form_response = _handle_drive_confirm_form(session, html_text, headers=auth_headers, job_id=job_id)
                if form_response and form_response.status_code == 200 and "text/html" not in form_response.headers.get("Content-Type", "").lower():
                    return _save_stream_to_file(form_response, dest_path, job_id=job_id)
            else:
                return _save_stream_to_file(response, dest_path, job_id=job_id)
    except Exception as e:
        log.warning(f"Direct URL download failed: {e}")

    # Method 4: Fall back to public download handler
    log.info(f"Falling back to public Drive download handler for '{file_id}'...")
    if job_id:
        log_job_message(job_id, f"🔄 Trying public Google Drive download handler...")
    return download_drive_file(file_id, dest_path, job_id=job_id)

def download_drive_file(file_id: str, dest_path: Path, job_id: Optional[str] = None) -> Path:
    """Download a Google Drive file by ID with stream chunking (32KB)."""
    session = requests.Session()
    direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    log.info(f"Stream downloading Drive file '{file_id}' → '{dest_path.name}'")
    if job_id:
        log_job_message(job_id, f"📥 Downloading Drive file: https://drive.google.com/uc?export=download&id={file_id}")
    response = session.get(direct_url, stream=True, timeout=120, allow_redirects=True)

    # Handle download_warning cookie
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break

    if token:
        confirm_url = f"https://drive.google.com/uc?export=download&confirm={token}&id={file_id}"
        response = session.get(confirm_url, stream=True, timeout=120, allow_redirects=True)

    # Handle HTML confirmation page or native Google Sheet export
    content_type = response.headers.get("Content-Type", "").lower()
    if response.status_code == 200 and "text/html" in content_type:
        html_text = response.text
        form_response = _handle_drive_confirm_form(session, html_text, job_id=job_id)
        if form_response and form_response.status_code == 200 and "text/html" not in form_response.headers.get("Content-Type", "").lower():
            return _save_stream_to_file(form_response, dest_path, job_id=job_id)

    if response.status_code != 200:
        err = f"Failed to download file from Google Drive (HTTP {response.status_code})"
        if job_id:
            log_job_message(job_id, f"✖ {err}", "ERROR")
        raise HTTPException(
            status_code=502,
            detail=err
        )

    return _save_stream_to_file(response, dest_path, job_id=job_id)

def _validate_downloaded_file(dest_path: Path) -> None:
    """Raise error if downloaded content is an HTML error page instead of a valid spreadsheet."""
    file_size = dest_path.stat().st_size
    if file_size < 100:
        raise HTTPException(
            status_code=400,
            detail=f"Downloaded file is too small ({file_size} bytes). File may be empty or inaccessible."
        )

    with open(dest_path, "rb") as f:
        head_bytes = f.read(4096)

    # Check for HTML error or login page
    head_text = head_bytes[:2048].decode("utf-8", errors="ignore").lower()
    if "<html" in head_text or "<!doctype html" in head_text or "accounts.google.com" in head_text or "sign in" in head_text:
        raise HTTPException(
            status_code=400,
            detail="Google Drive returned an HTML sign-in/error page instead of an Excel file. Check file sharing permissions ('Anyone with link' viewer) or pass an OAuth access_token."
        )
