import re
import logging
from pathlib import Path
from urllib.parse import urlparse, parse_qs
import requests
from fastapi import HTTPException
from typing import Optional

from core.logger import log_job_message

log = logging.getLogger("ei_stream_server.downloader")

# ─────────────────────────────────────────────────────────────────────────────
# Google Drive HTML confirmation form handler
# ─────────────────────────────────────────────────────────────────────────────

def _log(job_id: Optional[str], msg: str, level: str = "INFO"):
    """Log to both Python logger and per-job log buffer."""
    if level == "ERROR":
        log.error(msg)
    elif level == "WARNING":
        log.warning(msg)
    else:
        log.info(msg)
    if job_id:
        log_job_message(job_id, msg, level)


def _inspect_response(resp: requests.Response, label: str, job_id: Optional[str] = None) -> bool:
    """Check if a response contains actual file content (not HTML). Returns True if valid binary."""
    ct = resp.headers.get("Content-Type", "")
    cl = resp.headers.get("Content-Length", "?")
    _log(job_id, f"  [{label}] HTTP {resp.status_code} | Content-Type: {ct} | Content-Length: {cl}")

    if resp.status_code != 200:
        return False
    if "text/html" in ct.lower():
        # Peek at first 300 chars for debugging
        try:
            peek = resp.text[:300].replace("\n", " ").replace("\r", "")
        except Exception:
            peek = "(could not read body)"
        _log(job_id, f"  [{label}] ⚠ Got HTML page instead of binary. Peek: {peek[:200]}...", "WARNING")
        return False
    return True


def _handle_drive_confirm_form(session: requests.Session, html_text: str,
                                headers: dict = None, job_id: Optional[str] = None) -> Optional[requests.Response]:
    """Parse Google Drive HTML virus-scan / confirmation page and attempt to get the real file."""
    _log(job_id, "🔍 Parsing Google Drive HTML confirmation page...")

    # Extract file ID from the HTML for fallback URLs
    file_id_match = (re.search(r'name="id"\s+value="([^"]+)"', html_text) or
                     re.search(r'[?&]id=([a-zA-Z0-9_-]{25,})', html_text))
    fid = file_id_match.group(1) if file_id_match else None

    # Extract confirm token
    confirm_match = (re.search(r'name="confirm"\s+value="([^"]+)"', html_text) or
                     re.search(r'confirm=([a-zA-Z0-9_-]+)', html_text))
    confirm_val = confirm_match.group(1) if confirm_match else "t"

    # Extract uuid
    uuid_match = re.search(r'name="uuid"\s+value="([^"]+)"', html_text)
    uuid_val = uuid_match.group(1) if uuid_match else None

    _log(job_id, f"  Extracted: file_id={fid}, confirm={confirm_val}, uuid={uuid_val}")

    # Strategy 1: Parse <form> action and submit it
    action_match = re.search(r'<form[^>]*action="([^"]+)"', html_text)
    if action_match:
        action_url = action_match.group(1).replace("&amp;", "&")
        # Resolve relative URLs
        if action_url.startswith("//"):
            action_url = "https:" + action_url
        elif action_url.startswith("/"):
            action_url = "https://drive.usercontent.google.com" + action_url
        elif not action_url.startswith("http"):
            action_url = "https://drive.usercontent.google.com/" + action_url

        inputs = re.findall(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"', html_text)
        params = {name: val for name, val in inputs}

        _log(job_id, f"  📝 Strategy 1: Submitting form to {action_url} (params: {list(params.keys())})")
        try:
            resp = session.get(action_url, params=params, headers=headers or {},
                              stream=True, timeout=300, allow_redirects=True)
            if _inspect_response(resp, "form-submit", job_id):
                _log(job_id, "  ✅ Strategy 1 succeeded!")
                return resp
        except Exception as e:
            _log(job_id, f"  Strategy 1 failed: {e}", "WARNING")

    # Strategy 2: drive.usercontent.google.com direct download URL
    if fid:
        usercontent_urls = [
            f"https://drive.usercontent.google.com/download?id={fid}&export=download&confirm={confirm_val}" +
            (f"&uuid={uuid_val}" if uuid_val else ""),
        ]
        for i, u in enumerate(usercontent_urls):
            _log(job_id, f"  📝 Strategy 2.{i+1}: Trying usercontent URL: {u[:100]}...")
            try:
                resp = session.get(u, headers=headers or {}, stream=True, timeout=300, allow_redirects=True)
                if _inspect_response(resp, f"usercontent-{i+1}", job_id):
                    _log(job_id, f"  ✅ Strategy 2.{i+1} succeeded!")
                    return resp
            except Exception as e:
                _log(job_id, f"  Strategy 2.{i+1} failed: {e}", "WARNING")

    # Strategy 3: Classic uc confirm link
    if fid:
        classic_url = f"https://drive.google.com/uc?export=download&confirm={confirm_val}&id={fid}"
        _log(job_id, f"  📝 Strategy 3: Trying classic confirm URL: {classic_url[:100]}...")
        try:
            resp = session.get(classic_url, headers=headers or {}, stream=True, timeout=300, allow_redirects=True)
            if _inspect_response(resp, "classic-confirm", job_id):
                _log(job_id, "  ✅ Strategy 3 succeeded!")
                return resp
        except Exception as e:
            _log(job_id, f"  Strategy 3 failed: {e}", "WARNING")

    # Strategy 4: Follow any href containing 'download' or 'confirm'
    link_match = re.search(r'href="([^"]*(?:confirm|download)[^"]*)"', html_text)
    if link_match:
        link_url = link_match.group(1).replace("&amp;", "&")
        if link_url.startswith("//"):
            link_url = "https:" + link_url
        elif link_url.startswith("/"):
            link_url = "https://drive.google.com" + link_url
        elif not link_url.startswith("http"):
            link_url = "https://drive.google.com/" + link_url
        _log(job_id, f"  📝 Strategy 4: Following href link: {link_url[:100]}...")
        try:
            resp = session.get(link_url, headers=headers or {}, stream=True, timeout=300, allow_redirects=True)
            if _inspect_response(resp, "href-link", job_id):
                _log(job_id, "  ✅ Strategy 4 succeeded!")
                return resp
        except Exception as e:
            _log(job_id, f"  Strategy 4 failed: {e}", "WARNING")

    _log(job_id, "  ✖ All confirmation strategies exhausted. Could not bypass Google Drive HTML page.", "ERROR")
    return None


# ─────────────────────────────────────────────────────────────────────────────
# URL parsing helpers
# ─────────────────────────────────────────────────────────────────────────────

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


def _extract_token_from_url(url: str) -> Optional[str]:
    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    if "access_token" in qs:
        return qs["access_token"][0]
    return None


# ─────────────────────────────────────────────────────────────────────────────
# Stream-to-disk writer
# ─────────────────────────────────────────────────────────────────────────────

def _save_stream_to_file(response: requests.Response, dest_path: Path, job_id: Optional[str] = None) -> Path:
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    total_bytes = 0
    with open(dest_path, "wb") as f:
        for chunk in response.iter_content(chunk_size=32768):
            if chunk:
                f.write(chunk)
                total_bytes += len(chunk)
    _validate_downloaded_file(dest_path)
    size_mb = dest_path.stat().st_size / 1024 / 1024
    _log(job_id, f"💾 Saved to disk: '{dest_path.name}' ({size_mb:.2f} MB)")
    return dest_path


# ─────────────────────────────────────────────────────────────────────────────
# SINGLE ENTRY POINT: download_from_url
# ─────────────────────────────────────────────────────────────────────────────

def download_from_url(url: str, dest_path: Path, job_id: Optional[str] = None) -> Path:
    """
    Download a Google Drive file to disk. This is the SINGLE download entry point.
    Works with or without an OAuth access_token in the URL.

    Attempts in order:
      1. [If token] Drive API v3 Files.export (Google Sheets → xlsx)
      2. [If token] Drive API v3 Files.get alt=media (binary files)
      3. Direct GET on the supplied URL (handles redirects + HTML confirmation)
      4. Public uc?export=download with cookie + confirmation form handling
    """
    token = _extract_token_from_url(url)
    file_id = extract_file_id(url)

    _log(job_id, f"📥 Starting download | File ID: {file_id} | Has OAuth Token: {bool(token)} | URL: {url[:100]}...")

    session = requests.Session()
    auth_headers = {"Authorization": f"Bearer {token}"} if token else {}

    # ── Method 1: Drive API v3 Files.export (Google Sheets only) ─────────────
    if token:
        export_url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        _log(job_id, f"🔄 Method 1: Trying Drive API v3 export (Google Sheet)...")
        try:
            resp = session.get(export_url, headers=auth_headers, stream=True, timeout=120)
            if _inspect_response(resp, "api-export", job_id):
                return _save_stream_to_file(resp, dest_path, job_id=job_id)
            _log(job_id, f"  Method 1 returned HTTP {resp.status_code}, trying next...")
        except Exception as e:
            _log(job_id, f"  Method 1 failed: {e}", "WARNING")

    # ── Method 2: Drive API v3 alt=media (binary files) ──────────────────────
    if token:
        media_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
        _log(job_id, f"🔄 Method 2: Trying Drive API v3 alt=media (binary)...")
        try:
            resp = session.get(media_url, headers=auth_headers, stream=True, timeout=120)
            if _inspect_response(resp, "api-media", job_id):
                return _save_stream_to_file(resp, dest_path, job_id=job_id)
            _log(job_id, f"  Method 2 returned HTTP {resp.status_code}, trying next...")
        except Exception as e:
            _log(job_id, f"  Method 2 failed: {e}", "WARNING")

    # ── Method 3: Direct GET on supplied URL ─────────────────────────────────
    _log(job_id, f"🔄 Method 3: Direct GET on supplied URL...")
    try:
        resp = session.get(url, headers=auth_headers, stream=True, timeout=180, allow_redirects=True)
        if resp.status_code == 200:
            ct = resp.headers.get("Content-Type", "").lower()
            if "text/html" not in ct:
                _log(job_id, f"  ✅ Method 3: Got binary response (Content-Type: {ct})")
                return _save_stream_to_file(resp, dest_path, job_id=job_id)
            else:
                _log(job_id, f"  Method 3: Got HTML page, attempting confirmation form bypass...")
                html_text = resp.text
                form_resp = _handle_drive_confirm_form(session, html_text, headers=auth_headers, job_id=job_id)
                if form_resp:
                    return _save_stream_to_file(form_resp, dest_path, job_id=job_id)
        else:
            _log(job_id, f"  Method 3: HTTP {resp.status_code}, trying next...", "WARNING")
    except Exception as e:
        _log(job_id, f"  Method 3 failed: {e}", "WARNING")

    # ── Method 4: Public uc?export=download with cookie handling ─────────────
    public_url = f"https://drive.google.com/uc?export=download&id={file_id}"
    _log(job_id, f"🔄 Method 4: Public download with cookie handling...")
    try:
        resp = session.get(public_url, stream=True, timeout=180, allow_redirects=True)

        # Check for download_warning cookie (virus scan for large files)
        warn_token = None
        for key, value in resp.cookies.items():
            if key.startswith('download_warning'):
                warn_token = value
                break

        if warn_token:
            _log(job_id, f"  Found download_warning cookie: {warn_token}")
            confirm_url = f"https://drive.google.com/uc?export=download&confirm={warn_token}&id={file_id}"
            resp = session.get(confirm_url, stream=True, timeout=180, allow_redirects=True)

        ct = resp.headers.get("Content-Type", "").lower()
        if resp.status_code == 200 and "text/html" not in ct:
            _log(job_id, f"  ✅ Method 4: Got binary response")
            return _save_stream_to_file(resp, dest_path, job_id=job_id)

        if resp.status_code == 200 and "text/html" in ct:
            _log(job_id, f"  Method 4: Got HTML page, attempting confirmation form bypass...")
            html_text = resp.text
            form_resp = _handle_drive_confirm_form(session, html_text, job_id=job_id)
            if form_resp:
                return _save_stream_to_file(form_resp, dest_path, job_id=job_id)

        _log(job_id, f"  Method 4: HTTP {resp.status_code}, Content-Type: {ct}", "WARNING")
    except Exception as e:
        _log(job_id, f"  Method 4 failed: {e}", "WARNING")

    # ── All methods exhausted ────────────────────────────────────────────────
    _log(job_id, "✖ ALL 4 DOWNLOAD METHODS FAILED. File cannot be retrieved.", "ERROR")
    raise HTTPException(
        status_code=400,
        detail=(
            "Google Drive returned an HTML sign-in/error page instead of an Excel file. "
            "Check file sharing permissions ('Anyone with link' viewer) or pass an OAuth access_token."
        )
    )


# ─────────────────────────────────────────────────────────────────────────────
# File validation
# ─────────────────────────────────────────────────────────────────────────────

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

    head_text = head_bytes[:2048].decode("utf-8", errors="ignore").lower()
    if "<html" in head_text or "<!doctype html" in head_text or "accounts.google.com" in head_text or "sign in" in head_text:
        raise HTTPException(
            status_code=400,
            detail="Google Drive returned an HTML sign-in/error page instead of an Excel file. "
                   "Check file sharing permissions ('Anyone with link' viewer) or pass an OAuth access_token."
        )
