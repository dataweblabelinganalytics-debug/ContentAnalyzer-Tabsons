"""MongoDB-backed background worker for comparison processing jobs."""

import io
import logging
import os
import socket
import time
import traceback
import uuid
from datetime import datetime, timedelta

from gridfs import NoFile
from pymongo import ASCENDING, ReturnDocument
from pymongo.errors import PyMongoError

from app import (
    JOB_LEASE_SECONDS,
    JOB_UPLOAD_RETENTION_DAYS,
    get_job_files_bucket,
    get_mongo_database,
    get_processing_jobs_collection,
    iso_after,
    read_job_file,
    run_comparison,
    update_job_progress,
    upload_to_db,
    utc_now_iso,
)

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
POLL_SECONDS = float(os.getenv("JOB_POLL_SECONDS", "2"))
CLEANUP_INTERVAL_SECONDS = int(os.getenv("JOB_CLEANUP_INTERVAL_SECONDS", "3600"))
WORKER_ID = os.getenv(
    "JOB_WORKER_ID",
    f"{socket.gethostname()}-{os.getpid()}-{uuid.uuid4().hex[:8]}",
)

logging.basicConfig(
    level=getattr(logging, LOG_LEVEL, logging.INFO),
    format="%(asctime)s %(levelname)s [%(name)s] %(message)s",
)
logger = logging.getLogger("job-worker")


def recover_expired_jobs():
    now = utc_now_iso()
    result = get_processing_jobs_collection().update_many(
        {
            "status": "PROCESSING",
            "lease_expires_at": {"$lt": now},
        },
        {
            "$set": {
                "status": "QUEUED",
                "progress_percentage": 0,
                "current_step": "Recovered after worker interruption",
                "updated_at": now,
                "worker_id": None,
                "lease_expires_at": None,
            }
        },
    )
    if result.modified_count:
        logger.warning("Recovered %s abandoned processing jobs", result.modified_count)


def claim_next_job():
    now = utc_now_iso()
    return get_processing_jobs_collection().find_one_and_update(
        {"status": "QUEUED"},
        {
            "$set": {
                "status": "PROCESSING",
                "started_at": now,
                "completed_at": None,
                "updated_at": now,
                "progress_percentage": 5,
                "current_step": "Worker Started",
                "error_message": None,
                "worker_id": WORKER_ID,
                "lease_expires_at": iso_after(seconds=JOB_LEASE_SECONDS),
            },
            "$inc": {"attempt_count": 1},
        },
        sort=[("created_at", ASCENDING)],
        return_document=ReturnDocument.AFTER,
    )


def store_job_result(job_id, output_filename, xlsx_bytes):
    return get_job_files_bucket().upload_from_stream(
        output_filename,
        io.BytesIO(xlsx_bytes),
        metadata={
            "job_id": job_id,
            "kind": "result",
            "created_at": utc_now_iso(),
        },
    )


def process_job(job):
    job_id = job["job_id"]
    result_file_id = None
    logger.info(
        "Processing job job_id=%s filename=%s attempt=%s",
        job_id,
        job.get("original_filename"),
        job.get("attempt_count"),
    )

    def progress(progress_percentage, current_step):
        update_job_progress(
            job_id,
            progress_percentage,
            current_step,
            worker_id=WORKER_ID,
        )

    try:
        progress(10, "Reading Workbook")
        file_bytes = read_job_file(job["input_file_id"])
        xlsx_bytes, output_filename, stats = run_comparison(
            file_bytes,
            job.get("original_filename") or "upload.xlsx",
            progress_callback=progress,
        )

        channel = stats.get("channel", "UNKNOWN")
        date_str = stats.get("date", "00/00/0000")
        progress(85, "Uploading To MongoDB")
        mongo_timings = {}
        report_file_id = upload_to_db(
            xlsx_bytes,
            channel,
            date_str,
            output_filename,
            timing_stats=mongo_timings,
            request_id=job_id,
        )

        progress(95, "Finalizing Report")
        result_file_id = store_job_result(job_id, output_filename, xlsx_bytes)
        completed_at = utc_now_iso()
        timings = dict(stats.get("timings_ms") or {})
        timings.update(mongo_timings)
        completion_result = get_processing_jobs_collection().update_one(
            {
                "job_id": job_id,
                "status": "PROCESSING",
                "worker_id": WORKER_ID,
            },
            {
                "$set": {
                    "status": "COMPLETED",
                    "progress_percentage": 100,
                    "current_step": "Completed",
                    "completed_at": completed_at,
                    "updated_at": completed_at,
                    "error_message": None,
                    "output_filename": output_filename,
                    "result_file_id": result_file_id,
                    "result_size_bytes": len(xlsx_bytes),
                    "report_file_id": report_file_id,
                    "channel": channel,
                    "date": date_str,
                    "timings_ms": timings,
                    "worker_id": None,
                    "lease_expires_at": None,
                }
            },
        )
        if completion_result.modified_count != 1:
            get_job_files_bucket().delete(result_file_id)
            raise RuntimeError("Job lease was lost before completion")
        logger.info("Completed job job_id=%s report_file_id=%s", job_id, report_file_id)
    except Exception as exc:
        if result_file_id is not None:
            try:
                get_job_files_bucket().delete(result_file_id)
            except (NoFile, PyMongoError):
                logger.exception("Unable to remove failed result artifact job_id=%s", job_id)

        failed_at = utc_now_iso()
        error_message = f"{type(exc).__name__}: {exc}"[:4000]
        get_processing_jobs_collection().update_one(
            {"job_id": job_id, "worker_id": WORKER_ID},
            {
                "$set": {
                    "status": "FAILED",
                    "current_step": "Processing Failed",
                    "completed_at": failed_at,
                    "updated_at": failed_at,
                    "error_message": error_message,
                    "worker_id": None,
                    "lease_expires_at": None,
                }
            },
        )
        logger.error(
            "Job failed job_id=%s error=%s\n%s",
            job_id,
            error_message,
            traceback.format_exc(),
        )


def cleanup_expired_job_data():
    jobs = get_processing_jobs_collection()
    bucket = get_job_files_bucket()
    now = utc_now_iso()
    cleaned = 0

    expired_jobs = jobs.find(
        {
            "status": {"$in": ["COMPLETED", "FAILED"]},
            "input_file_id": {"$ne": None},
            "input_expires_at": {"$lte": now},
        },
        {"job_id": 1, "input_file_id": 1},
    )
    for job in expired_jobs:
        try:
            bucket.delete(job["input_file_id"])
        except NoFile:
            pass
        jobs.update_one(
            {"_id": job["_id"]},
            {
                "$unset": {"input_file_id": "", "input_expires_at": ""},
                "$set": {"input_cleaned_at": now, "updated_at": now},
            },
        )
        cleaned += 1

    orphan_cutoff = datetime.utcnow() - timedelta(days=JOB_UPLOAD_RETENTION_DAYS)
    files = get_mongo_database().job_files.files.find(
        {"uploadDate": {"$lte": orphan_cutoff}},
        {"_id": 1, "metadata.job_id": 1, "metadata.kind": 1},
    )
    orphaned = 0
    for file_document in files:
        metadata = file_document.get("metadata") or {}
        job_id = metadata.get("job_id")
        kind = metadata.get("kind")
        job = jobs.find_one(
            {"job_id": job_id},
            {"input_file_id": 1, "result_file_id": 1},
        ) if job_id else None
        referenced_file_id = (
            job.get("input_file_id") if job and kind == "input"
            else job.get("result_file_id") if job and kind == "result"
            else None
        )
        if not job or referenced_file_id != file_document["_id"]:
            try:
                bucket.delete(file_document["_id"])
                orphaned += 1
            except NoFile:
                pass

    if cleaned or orphaned:
        logger.info(
            "Cleanup complete expired_inputs=%s orphaned_files=%s",
            cleaned,
            orphaned,
        )


def run_worker():
    logger.info("Worker starting worker_id=%s poll_seconds=%s", WORKER_ID, POLL_SECONDS)
    recover_expired_jobs()
    next_cleanup = 0.0

    while True:
        try:
            if time.monotonic() >= next_cleanup:
                cleanup_expired_job_data()
                next_cleanup = time.monotonic() + CLEANUP_INTERVAL_SECONDS

            recover_expired_jobs()
            job = claim_next_job()
            if job:
                process_job(job)
                continue
        except PyMongoError:
            logger.exception("MongoDB worker loop error")
        except Exception:
            logger.exception("Unexpected worker loop error")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        run_worker()
    except KeyboardInterrupt:
        logger.info("Worker stopped")
