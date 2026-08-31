import logging

log = logging.getLogger("ei_stream_server")

def print_job_start(job_id: str, report_type: str, file_id: str):
    log.info(f"▶▶ START JOB [{job_id}] | Type: {report_type} | Target: {file_id}")

def print_job_step(job_id: str, step_num: int, message: str):
    log.info(f"  [{job_id}] Step {step_num}: {message}")

def print_job_success(job_id: str, duration_sec: float, output_filename: str):
    log.info(f"✔✔ SUCCESS JOB [{job_id}] | Duration: {duration_sec:.2f}s | Output: {output_filename}")

def print_job_error(job_id: str, error_msg: str):
    log.error(f"✖✖ FAILED JOB [{job_id}] | Error: {error_msg}")
