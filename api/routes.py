import os
import logging
from pathlib import Path
from typing import Optional
from fastapi import APIRouter, Depends, Query, File, UploadFile, HTTPException, Request
from fastapi.responses import FileResponse

from config.settings import CACHE_DIR
from core.auth import verify_api_key
from core.jobs import create_report_job, create_upload_report_job, get_job, active_jobs

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
    x_api_key: Optional[str] = Depends(verify_api_key)
):
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
    return {
        "job_id": job_id,
        "status": job.get("status"),
        "progress": job.get("progress"),
        "error": job.get("error"),
        "report_type": job.get("report_type"),
        "sub_type": job.get("sub_type"),
        "file_name": job.get("file_name"),
        "created_at": job.get("created_at"),
        "completed_at": job.get("completed_at")
    }

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

    download_name = job.get("file_name") or f"EI_REPORT_{job_id}.xlsx"
    return FileResponse(
        path=str(out_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name
    )
