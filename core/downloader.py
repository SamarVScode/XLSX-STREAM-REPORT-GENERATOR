import re
import logging
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import requests
from fastapi import HTTPException
from typing import Optional

log = logging.getLogger("ei_stream_server.downloader")

def _handle_drive_confirm_form(session: requests.Session, html_text: str) -> Optional[requests.Response]:
    """Parse Google Drive HTML warning page form (<form id='download-form'>) and submit it."""
    action_match = re.search(r'<form[^>]*action="([^"]+)"', html_text)
    if not action_match:
        return None

    action_url = action_match.group(1).replace("&amp;", "&")
    inputs = re.findall(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"', html_text)
    params = {name: val for name, val in inputs}

    log.info(f"Submitting Google Drive HTML confirmation form to: {action_url}")
    try:
        resp = session.get(action_url, params=params, stream=True, timeout=180)
        return resp
    except Exception as e:
        log.warning(f"Error submitting Drive confirmation form: {e}")
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

def is_direct_download_url(url: str) -> bool:
    """Returns True if the URL carries an explicit OAuth access_token or API key or export format."""
    return "access_token" in url or "alt=media" in url or "exportFormat=xlsx" in url

def download_from_url(url: str, dest_path: Path) -> Path:
    """
    Download a file directly from a fully-formed URL in 32KB stream chunks directly to disk.
    Also handles HTML confirmation pages for large files.
    """
    log.info(f"Direct stream download → '{dest_path.name}' from: {url[:120]}...")
    session = requests.Session()
    response = session.get(url, stream=True, timeout=120, allow_redirects=True)

    if response.status_code == 200:
        content_type = response.headers.get("Content-Type", "").lower()
        if "text/html" in content_type:
            html_text = response.text
            form_response = _handle_drive_confirm_form(session, html_text)
            if form_response and form_response.status_code == 200 and "text/html" not in form_response.headers.get("Content-Type", "").lower():
                response = form_response
            else:
                confirm_match = re.search(r'confirm=([a-zA-Z0-9_-]+)', html_text) or re.search(r'href="([^"]*confirm[^"]*)"', html_text)
                if confirm_match:
                    confirm_link = confirm_match.group(1)
                    if not confirm_link.startswith("http"):
                        confirm_url = f"https://drive.google.com/uc?export=download&confirm={confirm_link}&id={extract_file_id(url)}"
                    else:
                        confirm_url = confirm_link.replace("&amp;", "&")
                    log.info(f"Extracted Google Drive HTML confirm link: {confirm_url[:120]}...")
                    response = session.get(confirm_url, stream=True, timeout=120)

    if response.status_code != 200:
        log.warning(f"Direct URL download returned HTTP {response.status_code}. Retrying via file_id...")
        file_id = extract_file_id(url)
        return download_drive_file(file_id, dest_path)

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)

    try:
        _validate_downloaded_file(dest_path)
    except HTTPException as e:
        log.warning(f"Validation failed for direct URL ({e.detail}). Retrying with download_drive_file...")
        file_id = extract_file_id(url)
        return download_drive_file(file_id, dest_path)

    log.info(f"Direct stream download complete: '{dest_path.name}' ({dest_path.stat().st_size} bytes)")
    return dest_path

def download_drive_file(file_id: str, dest_path: Path) -> Path:
    """Download a Google Drive file by ID with stream chunking (32KB)."""
    session = requests.Session()
    direct_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    log.info(f"Stream downloading Drive file '{file_id}' → '{dest_path.name}'")
    response = session.get(direct_url, stream=True, timeout=120)

    # Handle download_warning cookie
    token = None
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            token = value
            break

    if token:
        confirm_url = f"https://drive.google.com/uc?export=download&confirm={token}&id={file_id}"
        response = session.get(confirm_url, stream=True, timeout=120)

    # Handle HTML confirmation page or native Google Sheet export
    content_type = response.headers.get("Content-Type", "").lower()
    if response.status_code == 200 and "text/html" in content_type:
        html_text = response.text
        form_response = _handle_drive_confirm_form(session, html_text)
        if form_response and form_response.status_code == 200 and "text/html" not in form_response.headers.get("Content-Type", "").lower():
            response = form_response
        else:
            confirm_match = re.search(r'confirm=([a-zA-Z0-9_-]+)', html_text)
            if confirm_match:
                confirm_token = confirm_match.group(1)
                confirm_url = f"https://drive.google.com/uc?export=download&confirm={confirm_token}&id={file_id}"
                log.info(f"Extracted confirm token '{confirm_token}' from HTML body.")
                response = session.get(confirm_url, stream=True, timeout=120)
            elif "docs.google.com/spreadsheets" in html_text or "google-apps.spreadsheet" in html_text or "Google Sheets" in html_text:
                log.info(f"File '{file_id}' is a native Google Sheet. Exporting as .xlsx...")
                export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?exportFormat=xlsx"
                response = session.get(export_url, stream=True, timeout=120)

    if response.status_code != 200:
        raise HTTPException(
            status_code=502,
            detail=f"Failed to download file from Google Drive (HTTP {response.status_code})"
        )

    dest_path.parent.mkdir(parents=True, exist_ok=True)
    with open(dest_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)

    _validate_downloaded_file(dest_path)
    log.info(f"Downloaded Drive file '{file_id}' ({dest_path.stat().st_size} bytes)")
    return dest_path

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
