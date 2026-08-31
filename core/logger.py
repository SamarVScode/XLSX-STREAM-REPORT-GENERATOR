import logging
import time
from typing import Dict, List
from collections import defaultdict

log = logging.getLogger("ei_stream_server")

# In-memory log buffer for active jobs
_job_logs: Dict[str, List[str]] = defaultdict(list)

def _append_job_log(job_id: str, message: str, level: str = "INFO"):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    log_line = f"[{timestamp}] [{level}] {message}"
    _job_logs[job_id].append(log_line)
    if len(_job_logs[job_id]) > 500:
        _job_logs[job_id].pop(0)

def log_job_message(job_id: str, message: str, level: str = "INFO"):
    _append_job_log(job_id, message, level)
    if level == "ERROR":
        log.error(f"[{job_id}] {message}")
    elif level == "WARNING":
        log.warning(f"[{job_id}] {message}")
    else:
        log.info(f"[{job_id}] {message}")

def print_job_start(job_id: str, report_type: str, file_id: str):
    msg = f"▶▶ START JOB [{job_id}] | Type: {report_type} | Target: {file_id}"
    log_job_message(job_id, msg, "INFO")

def print_job_step(job_id: str, step_num: int, message: str, duration: float = None):
    dur_str = f" (took {duration:.2f}s)" if duration is not None else ""
    msg = f"Step {step_num}: {message}{dur_str}"
    log_job_message(job_id, msg, "INFO")

def print_job_success(job_id: str, duration_sec: float, output_filename: str):
    msg = f"✔✔ SUCCESS JOB [{job_id}] | Total Duration: {duration_sec:.2f}s | Output: {output_filename}"
    log_job_message(job_id, msg, "INFO")

def print_job_error(job_id: str, error_msg: str):
    msg = f"✖✖ FAILED JOB [{job_id}] | Error: {error_msg}"
    log_job_message(job_id, msg, "ERROR")

def get_job_log_lines(job_id: str) -> List[str]:
    return _job_logs.get(job_id, [])
