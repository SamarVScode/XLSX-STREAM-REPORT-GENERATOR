import uuid
import time
import json
import logging
import threading
import re
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional
from fastapi import HTTPException

from config.settings import CACHE_DIR, CACHE_TTL, CACHE_MAX_AGE, MAX_CONCURRENT_JOBS
from core.downloader import extract_file_id, download_from_url
from core.logger import print_job_start, print_job_step, print_job_success, print_job_error, log_job_message
from generators import (
    generate_ei_report,
    generate_forward_pendency_report,
    generate_conversion_report,
    extract_report_date_from_agent_view,
    generate_reverse_pendency_report,
    generate_nps_report,
    generate_tat_report,
    generate_vms_adherence_report,
    generate_second_attempt_adherence_report,
    generate_eob_report,
    generate_untraceable_report
)

log = logging.getLogger("ei_stream_server.jobs")

active_jobs: Dict[str, Dict[str, Any]] = {}
conversion_semaphore = threading.Semaphore(MAX_CONCURRENT_JOBS)
_jobs_lock = threading.Lock()


def _job_meta_path(job_id: str) -> Path:
    return CACHE_DIR / f"JOB_{job_id}.json"


def _save_job(job_id: str, job: Dict[str, Any]) -> None:
    try:
        meta_path = _job_meta_path(job_id)
        with open(meta_path, "w") as f:
            json.dump({"job_id": job_id, **job}, f)
    except Exception as e:
        log.warning(f"Could not persist job {job_id} to disk: {e}")


def _set_job(job_id: str, job: Dict[str, Any]) -> None:
    with _jobs_lock:
        active_jobs[job_id] = job
    _save_job(job_id, job)


def recover_jobs_from_disk() -> None:
    recovered = 0
    now = time.time()
    try:
        for meta_file in CACHE_DIR.glob("JOB_*.json"):
            try:
                with open(meta_file, "r") as f:
                    data = json.load(f)
                jid = data.get("job_id")
                created_at = data.get("created_at", 0)
                if not jid:
                    continue

                if now - created_at > CACHE_MAX_AGE:
                    meta_file.unlink(missing_ok=True)
                    out_p = data.get("output_path")
                    if out_p:
                        Path(out_p).unlink(missing_ok=True)
                    continue

                if data.get("status") == "processing":
                    data["status"] = "error"
                    data["error"] = "Job interrupted by server restart."
                    data["progress"] = "Failed (server restarted)"

                active_jobs[jid] = data
                recovered += 1
            except Exception as e:
                log.warning(f"Could not recover {meta_file.name}: {e}")
        if recovered:
            log.info(f"Recovered {recovered} jobs from disk on boot.")
    except Exception as e:
        log.warning(f"Error during job recovery: {e}")


def get_job(job_id: str) -> Dict[str, Any]:
    if job_id in active_jobs:
        return active_jobs[job_id]

    meta_file = _job_meta_path(job_id)
    if meta_file.exists():
        try:
            with open(meta_file, "r") as f:
                data = json.load(f)
            active_jobs[job_id] = data
            return data
        except Exception:
            pass

    raise HTTPException(status_code=404, detail="Job not found")


def cleanup_old_jobs() -> None:
    evict_old_jobs()


def evict_old_jobs() -> None:
    now = time.time()
    to_delete = []
    for jid, job in list(active_jobs.items()):
        created_at = job.get("created_at", 0)
        if now - created_at > CACHE_MAX_AGE:
            to_delete.append(jid)
    for jid in to_delete:
        meta_file = _job_meta_path(jid)
        meta_file.unlink(missing_ok=True)
        job = active_jobs.get(jid, {})
        out = job.get("output_path")
        if out:
            Path(out).unlink(missing_ok=True)
        del active_jobs[jid]
        log.info(f"Evicted expired job {jid} from cache")


def generate_proper_report_filename(
    report_type: str,
    original_filename: Optional[str] = None,
    sub_type: Optional[str] = None,
    report_date: Optional[str] = None
) -> str:
    prefix_map = {
        "ei": "EI_Summary_Report",
        "forward_pendency": "Forward_Pendency_Report",
        "reverse_pendency": "Reverse_Pendency_Report",
        "conversion": "Conversion_Summary_Report",
        "nps": "NPS_Report",
        "tat": "SCM_TAT_24Hrs_Performance_Report",
        "scm_tat": "SCM_TAT_24Hrs_Performance_Report",
        "scm-tat": "SCM_TAT_24Hrs_Performance_Report",
        "vms": "VMS_Adherence_Report",
        "vms_adherence": "VMS_Adherence_Report",
        "vms-adherence": "VMS_Adherence_Report",
        "second_attempt_adherence": "2nd_Attempt_Adherence_Report",
        "2nd_attempt_adherence": "2nd_Attempt_Adherence_Report",
        "2nd_adherence": "2nd_Attempt_Adherence_Report",
        "adherence_2nd": "2nd_Attempt_Adherence_Report",
        "2nd_attempt": "2nd_Attempt_Adherence_Report",
        "eob": "EOB_Priority_Report",
        "eob_report": "EOB_Priority_Report",
        "eob-report": "EOB_Priority_Report",
        "untraceable": "Untraceable_Report",
        "untraceable_report": "Untraceable_Report",
        "untraceable-report": "Untraceable_Report",
        "ut": "Untraceable_Report"
    }

    norm_type = report_type.lower().strip().replace(" ", "_")
    prefix = prefix_map.get(norm_type, f"{report_type.upper()}_Report")

    if norm_type == "conversion":
        if sub_type:
            clean_sub = sub_type.strip().lower()
            if clean_sub in ("d1", "d-1", "nextday"):
                prefix = "Conversion_Summary_D-1_Report"
            elif clean_sub == "sameday":
                prefix = "Conversion_Summary_Sameday_Report"
            else:
                prefix = f"Conversion_Summary_{sub_type.strip().replace(' ', '_')}_Report"
        else:
            prefix = "Conversion_Summary_Report"

        gen_date_str = datetime.now().strftime("%d-%b-%Y")
        if report_date and str(report_date).strip():
            clean_report_date = str(report_date).strip().replace('/', '-').replace('_', '-')
            return f"{prefix}_{clean_report_date}_{gen_date_str}.xlsx"
        elif original_filename:
            date_match = re.search(r'(\d{1,2}[-_][a-zA-Z]{3}[-_]\d{4}|\d{4}[-_]\d{2}[-_]\d{2}|\d{1,2}[-_]\d{1,2}[-_]\d{4})', original_filename)
            if date_match:
                extracted_date = date_match.group(1).replace('_', '-')
                return f"{prefix}_{extracted_date}_{gen_date_str}.xlsx"
        return f"{prefix}_{gen_date_str}.xlsx"

    date_str = datetime.now().strftime("%d-%b-%Y")
    if original_filename:
        date_match = re.search(r'(\d{1,2}[-_][a-zA-Z]{3}[-_]\d{4}|\d{4}[-_]\d{2}[-_]\d{2})', original_filename)
        if date_match:
            extracted_date = date_match.group(1).replace('_', '-')
            return f"{prefix}_{extracted_date}.xlsx"

    return f"{prefix}_{date_str}.xlsx"


def background_report_job(job_id: str, file_id: str, output_path: Path, report_type: str = "ei") -> None:
    log_job_message(job_id, f"Waiting for execution lock (Max concurrent: {MAX_CONCURRENT_JOBS})...")
    acquired = conversion_semaphore.acquire(timeout=600)
    if not acquired:
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = "Server busy: concurrent limit reached."
        job["progress"] = "Failed: server busy"
        _set_job(job_id, job)
        print_job_error(job_id, "Server busy: concurrent limit reached")
        return

    # Redundancy check
    if output_path.exists() and output_path.stat().st_size > 0:
        conversion_semaphore.release()
        job = active_jobs.get(job_id, {})
        job["status"] = "done"
        job["progress"] = "Complete (cached)"
        job["completed_at"] = output_path.stat().st_mtime
        _set_job(job_id, job)
        log_job_message(job_id, f"Report already cached ({output_path.name}). Returning instant result.")
        return

    t_start = time.time()
    source_desc = "Direct Upload" if str(file_id).startswith("upload_") else f"Google Drive ({file_id})"
    print_job_start(job_id, report_type, source_desc)

    tmp_input = None
    try:
        job = active_jobs.get(job_id, {})
        job["status"] = "processing"
        job["progress"] = "Locating input file..."
        _set_job(job_id, job)

        for path in CACHE_DIR.glob(f"{file_id}.*"):
            if path.suffix.lower() not in ('.json', '.log') and not path.name.startswith("REPORT_"):
                tmp_input = path
                break

        if tmp_input is None:
            tmp_input = CACHE_DIR / f"{file_id}.xlsx"

        if not file_id.startswith("upload_"):
            needs_download = not tmp_input.exists() or (time.time() - tmp_input.stat().st_mtime > CACHE_TTL)
            if needs_download:
                t0 = time.time()
                raw_url = active_jobs.get(job_id, {}).get("source_url", "")
                download_url = raw_url if raw_url else f"https://drive.google.com/uc?export=download&id={file_id}"
                print_job_step(job_id, 1, f"Downloading source file (ID: {file_id})...")
                download_from_url(download_url, tmp_input, job_id=job_id)
                print_job_step(job_id, 1, f"Source file downloaded ({tmp_input.stat().st_size / 1024 / 1024:.2f} MB in {time.time() - t0:.2f}s)")
            else:
                print_job_step(job_id, 1, f"Using cached source file ({tmp_input.name})")

        job["progress"] = f"Generating {report_type} Stream Report..."
        _set_job(job_id, job)

        t_gen = time.time()
        print_job_step(job_id, 2, f"Executing stream report generator ({report_type})...")

        norm_type = report_type.lower().strip().replace("-", "_")

        if norm_type == "forward_pendency":
            generate_forward_pendency_report(tmp_input, output_path)
        elif norm_type == "conversion":
            sub_type = job.get("sub_type") or "sameday"
            generate_conversion_report(tmp_input, output_path, sub_type=sub_type)
            try:
                report_date = extract_report_date_from_agent_view(tmp_input)
                if not report_date and output_path.exists():
                    report_date = extract_report_date_from_agent_view(output_path)
            except Exception as ex:
                log.warning(f"Could not extract report date from Agent_view: {ex}")
                report_date = None
            if report_date:
                job["report_date"] = report_date
                job["file_name"] = generate_proper_report_filename(
                    "conversion",
                    original_filename=job.get("original_filename"),
                    sub_type=sub_type,
                    report_date=report_date
                )
        elif norm_type == "reverse_pendency":
            generate_reverse_pendency_report(tmp_input, output_path)
        elif norm_type == "nps":
            generate_nps_report(tmp_input, output_path)
        elif norm_type in ("tat", "scm_tat"):
            generate_tat_report(tmp_input, output_path)
        elif norm_type in ("vms", "vms_adherence"):
            generate_vms_adherence_report(tmp_input, output_path)
        elif norm_type in ("second_attempt_adherence", "2nd_attempt_adherence", "2nd_adherence", "adherence_2nd", "2nd_attempt"):
            generate_second_attempt_adherence_report(tmp_input, output_path)
        elif norm_type in ("eob", "eob_report"):
            generate_eob_report(tmp_input, output_path)
        elif norm_type in ("untraceable", "untraceable_report", "ut"):
            generate_untraceable_report(tmp_input, output_path)
        else:
            generate_ei_report(str(tmp_input), str(output_path))

        print_job_step(job_id, 2, f"Stream report generated in {time.time() - t_gen:.2f}s")

        # Immediate Disk Reclamation
        if tmp_input and tmp_input.exists():
            try:
                tmp_input.unlink()
                log_job_message(job_id, f"🗑️ Disk Reclaimed: Temporary raw input file deleted ({tmp_input.name})")
            except Exception as del_err:
                log.warning(f"Could not delete temporary input file {tmp_input.name}: {del_err}")

        job["status"] = "done"
        job["progress"] = "Complete"
        job["completed_at"] = time.time()
        _set_job(job_id, job)

        duration = time.time() - t_start
        print_job_success(job_id, duration, output_path.name)

    except Exception as e:
        log.exception(f"Job {job_id} failed: {e}")
        job = active_jobs.get(job_id, {})
        job["status"] = "error"
        job["error"] = str(e)
        job["progress"] = f"Failed: {str(e)[:100]}"
        _set_job(job_id, job)
        print_job_error(job_id, str(e))
        
        if tmp_input and tmp_input.exists():
            try:
                tmp_input.unlink()
            except Exception:
                pass
    finally:
        conversion_semaphore.release()


def create_job(drive_url: str, report_type: str = "ei", sub_type: Optional[str] = None) -> Dict[str, str]:
    return create_report_job(drive_url, report_type, sub_type)


def process_job_async(job_id: str, file_id: str, output_path: Path, report_type: str = "ei") -> None:
    thread = threading.Thread(
        target=background_report_job,
        args=(job_id, file_id, output_path, report_type),
        daemon=True
    )
    thread.start()


def create_report_job(drive_url: str, report_type: str = "ei", sub_type: Optional[str] = None) -> Dict[str, str]:
    evict_old_jobs()
    file_id = extract_file_id(drive_url)
    cache_key = f"{file_id}_{report_type}"
    output_path = CACHE_DIR / f"REPORT_{cache_key}.xlsx"
    formatted_name = generate_proper_report_filename(report_type, sub_type=sub_type)

    if output_path.exists() and (time.time() - output_path.stat().st_mtime <= CACHE_TTL):
        log_job_message(cache_key, f"Instant cache hit for {report_type} report.")
        cached_report_date = None
        if report_type == "conversion":
            try:
                cached_report_date = extract_report_date_from_agent_view(output_path)
            except Exception:
                pass
        formatted_name = generate_proper_report_filename(report_type, sub_type=sub_type, report_date=cached_report_date)
        active_jobs[cache_key] = {
            "status": "done",
            "output_path": str(output_path),
            "error": None,
            "progress": "Complete (cached)",
            "created_at": output_path.stat().st_mtime,
            "report_type": report_type,
            "sub_type": sub_type,
            "report_date": cached_report_date,
            "file_name": formatted_name
        }
        return {"job_id": cache_key, "status": "done"}

    job_id = cache_key
    job = {
        "status": "processing",
        "output_path": str(output_path),
        "error": None,
        "progress": "Job created, waiting to download...",
        "created_at": time.time(),
        "source_url": drive_url,
        "report_type": report_type,
        "sub_type": sub_type,
        "file_name": formatted_name
    }
    _set_job(job_id, job)
    log_job_message(job_id, f"Initialized async job for {report_type} (source: {drive_url[:80]}...)")

    process_job_async(job_id, file_id, output_path, report_type)
    return {"job_id": job_id, "status": "processing"}


def create_upload_report_job(file_bytes: bytes, filename: str = "upload.xlsx", report_type: str = "ei", sub_type: Optional[str] = None) -> Dict[str, str]:
    evict_old_jobs()
    job_id = str(uuid.uuid4())[:8]
    file_id = f"upload_{job_id}"
    log.info(f"Creating upload stream job {job_id} (filename={filename}, size={len(file_bytes)} bytes, type={report_type})")

    ext = Path(filename).suffix.lower() if filename and Path(filename).suffix else ".xlsx"
    tmp_input = CACHE_DIR / f"{file_id}{ext}"
    with open(tmp_input, "wb") as f:
        f.write(file_bytes)

    output_path = CACHE_DIR / f"REPORT_{job_id}.xlsx"
    formatted_name = generate_proper_report_filename(report_type, original_filename=filename, sub_type=sub_type)
    job = {
        "status": "processing",
        "output_path": str(output_path),
        "error": None,
        "progress": "Uploaded, starting generation...",
        "created_at": time.time(),
        "report_type": report_type,
        "sub_type": sub_type,
        "original_filename": filename,
        "file_name": formatted_name
    }
    _set_job(job_id, job)
    log_job_message(job_id, f"Upload received: {filename} ({len(file_bytes)} bytes). Starting generator...")

    process_job_async(job_id, file_id, output_path, report_type)
    return {"job_id": job_id, "status": "processing"}
