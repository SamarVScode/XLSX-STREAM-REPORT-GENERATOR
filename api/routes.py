import os
import logging
from pathlib import Path
from typing import Optional
from fastapi import APIRouter, Depends, Query, File, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse

from config.settings import CACHE_DIR
from core.auth import verify_api_key
from core.jobs import create_report_job, create_upload_report_job, get_job, active_jobs
from core.logger import get_job_log_lines

router = APIRouter()
log = logging.getLogger("ei_stream_server.routes")

ROUTE_TYPE_MAP = {
    "/reports/ei": "ei",
    "/reports/forward-pendency": "forward_pendency",
    "/reports/reverse-pendency": "reverse_pendency",
    "/reports/conversion": "conversion",
    "/reports/nps": "nps",
    "/reports/tat": "tat",
    "/reports/scm-tat": "tat",
    "/reports/vms-adherence": "vms_adherence",
    "/reports/second-attempt-adherence": "2nd_attempt_adherence",
    "/reports/2nd-attempt-adherence": "2nd_attempt_adherence",
    "/reports/eob": "eob",
    "/reports/untraceable": "untraceable",
    "/reports/ut": "untraceable",
}

@router.get("/")
async def root():
    return {
        "server": "EI Stream Server",
        "status": "ready",
        "version": "3.0.0",
        "engine": "In-Flight Rust Calamine / XML Streaming"
    }

@router.get("/health")
async def health(x_api_key: Optional[str] = Depends(verify_api_key)):
    active_count = sum(1 for j in active_jobs.values() if j.get("status") == "processing")
    return {
        "status": "ok",
        "active_jobs": active_count,
        "total_tracked_jobs": len(active_jobs),
        "cache_dir": str(CACHE_DIR)
    }

@router.get("/convert-async")
async def convert_async(
    drive_url: str = Query(..., description="Google Drive URL"),
    report_type: str = Query("ei", description="Report generator type"),
    sub_type: Optional[str] = Query(None, description="Report sub-type (e.g. sameday, d1)"),
    access_token: Optional[str] = Query(None, description="Optional Google OAuth access token"),
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    if access_token and "access_token=" not in drive_url:
        sep = "&" if "?" in drive_url else "?"
        drive_url = f"{drive_url}{sep}access_token={access_token}"
    return create_report_job(drive_url, report_type=report_type, sub_type=sub_type)

@router.post("/convert-upload")
@router.post("/generate-report")
@router.post("/reports/ei")
@router.post("/reports/forward-pendency")
@router.post("/reports/reverse-pendency")
@router.post("/reports/conversion")
@router.post("/reports/nps")
@router.post("/reports/tat")
@router.post("/reports/scm-tat")
@router.post("/reports/vms-adherence")
@router.post("/reports/second-attempt-adherence")
@router.post("/reports/2nd-attempt-adherence")
@router.post("/reports/eob")
@router.post("/reports/untraceable")
@router.post("/reports/ut")
async def convert_upload(
    request: Request,
    file: UploadFile = File(...),
    report_type: Optional[str] = Query(None, description="Report generator type"),
    sub_type: Optional[str] = Query(None, description="Report sub-type (e.g. sameday, d1)"),
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    if not report_type:
        path = request.url.path
        report_type = ROUTE_TYPE_MAP.get(path, "ei")
    content = await file.read()
    return create_upload_report_job(content, file.filename or "upload.xlsx", report_type=report_type, sub_type=sub_type)

@router.post("/reports/async/{report_type}")
async def trigger_async_report(
    report_type: str,
    file: UploadFile = File(...),
    sub_type: Optional[str] = Query(None, description="Report sub-type (e.g. sameday, d1)"),
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    content = await file.read()
    return create_upload_report_job(content, file.filename or "upload.xlsx", report_type=report_type, sub_type=sub_type)

@router.get("/job/{job_id}")
@router.get("/jobs/{job_id}")
async def get_job_status_route(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    job = get_job(job_id)
    logs = get_job_log_lines(job_id)
    tabs = job.get("tabs")
    if not tabs:
        report_type = str(job.get("report_type") or "").lower().replace("-", "_")
        default_tabs = {
            "2nd_attempt_adherence": ["Summary", "Raw"],
            "second_attempt_adherence": ["Summary", "Raw"],
            "ei": ["SUMMARY", "Filtered_Source_DC", "FWD EI", "REVERSE EI", "Agent Summary"],
            "tat": ["Summary", "SCM TAT raw data"],
            "scm_tat": ["Summary", "SCM TAT raw data"],
            "forward_pendency": ["Summary", "Raw", "CPD-DID"],
            "reverse_pendency": ["Summary", "Raw", "P0"],
            "conversion": ["Summary", "Raw"],
            "nps": ["Summary", "Raw"],
            "vms_adherence": ["Summary", "Raw"],
            "vms": ["Summary", "Raw"],
            "eob": ["Summary", "Raw"],
            "untraceable": ["Summary", "Raw"],
            "ut": ["Summary", "Raw"]
        }
        tabs = default_tabs.get(report_type, ["Summary", "Raw"])

    return {
        "job_id": job_id,
        "status": job.get("status"),
        "progress": job.get("progress"),
        "error": job.get("error"),
        "report_type": job.get("report_type"),
        "sub_type": job.get("sub_type"),
        "report_date": job.get("report_date"),
        "file_name": job.get("file_name"),
        "tabs": tabs,
        "created_at": job.get("created_at"),
        "completed_at": job.get("completed_at"),
        "logs": logs
    }

@router.get("/job/{job_id}/logs")
@router.get("/jobs/{job_id}/logs")
async def get_job_logs_route(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    logs = get_job_log_lines(job_id)
    return {
        "job_id": job_id,
        "logs": logs
    }

import time
from starlette.background import BackgroundTask

def _cleanup_job_cache(job_id: str, out_path: Path):
    """Auto-clean memory and disk cache once the report is downloaded, matching clear_jobs.py."""
    try:
        time.sleep(3)
        if out_path.exists():
            out_path.unlink(missing_ok=True)
        meta_p = CACHE_DIR / f"JOB_{job_id}.json"
        if meta_p.exists():
            meta_p.unlink(missing_ok=True)
        # Clear any source files and temp outputs associated with this job
        for p in CACHE_DIR.glob(f"*{job_id}*"):
            p.unlink(missing_ok=True)
        for p in CACHE_DIR.glob(f"upload_{job_id}*"):
            p.unlink(missing_ok=True)
        # Remove from active_jobs memory map
        if job_id in active_jobs:
            del active_jobs[job_id]
        log.info(f"🧹 Cleaned up cache for job {job_id} after download.")
    except Exception as e:
        log.warning(f"Cleanup warning for job {job_id}: {e}")

@router.get("/job/{job_id}/result")
@router.get("/jobs/{job_id}/result")
@router.get("/reports/download/{job_id}")
async def get_job_result(
    job_id: str,
    x_api_key: Optional[str] = Depends(verify_api_key)
):
    job = get_job(job_id)
    if job.get("status") != "done":
        raise HTTPException(status_code=400, detail=f"Job is not ready yet (status: {job.get('status')})")

    out_path = Path(job.get("output_path", ""))
    if not out_path.exists():
        raise HTTPException(status_code=404, detail="Generated report file has expired or was removed from cache.")

    if out_path.stat().st_size < 1000:
        raise HTTPException(status_code=500, detail=f"Generated report file is incomplete or empty ({out_path.stat().st_size} bytes).")

    download_name = job.get("file_name") or f"EI_REPORT_{job_id}.xlsx"
    return FileResponse(
        path=str(out_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name
    )

@router.post("/clear-cache")
@router.get("/clear-cache")
async def clear_cache_endpoint(x_api_key: Optional[str] = Depends(verify_api_key)):
    """Wipes all cached jobs and temporary report files from disk (matching clear_jobs.py)."""
    patterns = ["JOB_*.json", "REPORT_*.xlsx", "*.xlsx", "*.xlsb", "*.ods", "*.csv", "*.tsv", "*.xls", "*.xlsm"]
    deleted = 0
    for pattern in patterns:
        for f in CACHE_DIR.glob(pattern):
            try:
                f.unlink(missing_ok=True)
                deleted += 1
            except Exception:
                pass
    active_jobs.clear()
    log.info(f"🧹 Full cache wipe executed: deleted {deleted} files.")
    return {"status": "ok", "deleted_files": deleted, "cache_dir": str(CACHE_DIR)}
