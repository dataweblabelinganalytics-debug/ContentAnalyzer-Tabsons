import io
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from bson.objectid import ObjectId

import app
import job_worker


def queued_job(job_id="job-1"):
    return {
        "job_id": job_id,
        "status": "QUEUED",
        "progress_percentage": 0,
        "current_step": "Waiting for worker",
        "created_at": "2026-06-11T10:00:00",
        "started_at": None,
        "completed_at": None,
        "error_message": None,
        "original_filename": "input.xlsx",
        "attempt_count": 0,
        "input_file_id": ObjectId(),
    }


class FakeJobs:
    def __init__(self, document=None):
        self.document = document or queued_job()
        self.updates = []

    def find_one(self, query, projection=None):
        return dict(self.document) if self.document else None

    def find_one_and_update(self, query, update, **kwargs):
        if not self.document:
            return None
        self.document.update(update.get("$set", {}))
        return dict(self.document)

    def update_one(self, query, update):
        self.updates.append((query, update))
        self.document.update(update.get("$set", {}))
        for key in update.get("$unset", {}):
            self.document.pop(key, None)
        return SimpleNamespace(modified_count=1)


class FakeBucket:
    def __init__(self):
        self.result_id = ObjectId()
        self.deleted = []

    def upload_from_stream(self, filename, source, metadata=None):
        return self.result_id

    def delete(self, file_id):
        self.deleted.append(file_id)


class JobApiTests(unittest.TestCase):
    def tearDown(self):
        app._embedded_worker_process = None

    def test_standalone_web_service_starts_worker_process(self):
        process = SimpleNamespace(pid=1234, poll=lambda: None)
        with (
            patch.object(app, "MONGO_URI", "mongodb://example"),
            patch.object(app, "ENABLE_EMBEDDED_JOB_WORKER", True),
            patch.dict("os.environ", {}, clear=False),
            patch.object(app.subprocess, "Popen", return_value=process) as popen,
        ):
            app._embedded_worker_process = None
            app.start_embedded_job_worker()

        self.assertIs(app._embedded_worker_process, process)
        worker_env = popen.call_args.kwargs["env"]
        self.assertEqual(worker_env["JOB_WORKER_PROCESS"], "embedded")
        self.assertIn("job_worker.py", popen.call_args.args[0][-1])

    def test_compare_returns_job_immediately(self):
        job = queued_job()
        with patch.object(app, "create_processing_job", return_value=job):
            with app.app.test_client() as client:
                response = client.post(
                    "/api/compare",
                    data={"file": (io.BytesIO(b"test"), "input.xlsx")},
                    content_type="multipart/form-data",
                )

        self.assertEqual(response.status_code, 202)
        payload = response.get_json()
        self.assertEqual(payload["job_id"], "job-1")
        self.assertEqual(payload["status"], "QUEUED")
        self.assertIn("/api/job-status/job-1", payload["status_url"])

    def test_status_exposes_required_fields(self):
        job = queued_job()
        jobs = FakeJobs(job)
        with patch.object(app, "get_processing_jobs_collection", return_value=jobs):
            with app.app.test_client() as client:
                response = client.get("/api/job-status/job-1")

        self.assertEqual(response.status_code, 200)
        payload = response.get_json()
        for field in (
            "status",
            "progress",
            "current_step",
            "error",
            "download_available",
        ):
            self.assertIn(field, payload)

    def test_retry_requeues_failed_job_without_removing_upload(self):
        job = queued_job()
        job.update(
            {
                "status": "FAILED",
                "error_message": "failure",
                "completed_at": "2026-06-11T10:01:00",
            }
        )
        input_file_id = job["input_file_id"]
        jobs = FakeJobs(job)
        with patch.object(app, "get_processing_jobs_collection", return_value=jobs):
            with app.app.test_client() as client:
                response = client.post("/api/jobs/job-1/retry")

        self.assertEqual(response.status_code, 202)
        self.assertEqual(jobs.document["status"], "QUEUED")
        self.assertEqual(jobs.document["input_file_id"], input_file_id)

    def test_completed_job_download(self):
        job = queued_job()
        job.update(
            {
                "status": "COMPLETED",
                "result_file_id": ObjectId(),
                "output_filename": "result.xlsx",
            }
        )
        jobs = FakeJobs(job)
        with (
            patch.object(app, "get_processing_jobs_collection", return_value=jobs),
            patch.object(app, "read_job_file", return_value=b"xlsx"),
        ):
            with app.app.test_client() as client:
                response = client.get("/api/jobs/job-1/download")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.data, b"xlsx")
        self.assertIn("result.xlsx", response.headers["Content-Disposition"])


class WorkerTests(unittest.TestCase):
    def test_worker_completes_job_and_preserves_report_contract(self):
        job = queued_job()
        job.update({"status": "PROCESSING", "worker_id": job_worker.WORKER_ID})
        jobs = FakeJobs(job)
        bucket = FakeBucket()
        progress_updates = []

        def fake_run(file_bytes, original_name, progress_callback=None):
            progress_callback(25, "Analyzing BARC Data")
            progress_callback(65, "Generating Workbook")
            return (
                b"generated-xlsx",
                "result.xlsx",
                {
                    "channel": "CHANNEL",
                    "date": "10/05/2026",
                    "timings_ms": {"comparison_processor": 100},
                },
            )

        with (
            patch.object(job_worker, "get_processing_jobs_collection", return_value=jobs),
            patch.object(job_worker, "get_job_files_bucket", return_value=bucket),
            patch.object(job_worker, "read_job_file", return_value=b"input-xlsx"),
            patch.object(job_worker, "run_comparison", side_effect=fake_run),
            patch.object(job_worker, "upload_to_db", return_value="report-file-id"),
            patch.object(
                job_worker,
                "update_job_progress",
                side_effect=lambda job_id, progress, step, worker_id=None: progress_updates.append(
                    (progress, step)
                ),
            ),
        ):
            job_worker.process_job(job)

        self.assertEqual(jobs.document["status"], "COMPLETED")
        self.assertEqual(jobs.document["progress_percentage"], 100)
        self.assertEqual(jobs.document["report_file_id"], "report-file-id")
        self.assertEqual(jobs.document["result_file_id"], bucket.result_id)
        self.assertIn((85, "Uploading To MongoDB"), progress_updates)

    def test_worker_failure_keeps_upload_for_retry(self):
        job = queued_job()
        job.update({"status": "PROCESSING", "worker_id": job_worker.WORKER_ID})
        jobs = FakeJobs(job)
        input_file_id = job["input_file_id"]

        with (
            patch.object(job_worker, "get_processing_jobs_collection", return_value=jobs),
            patch.object(job_worker, "read_job_file", side_effect=ValueError("bad workbook")),
            patch.object(job_worker, "update_job_progress"),
        ):
            job_worker.process_job(job)

        self.assertEqual(jobs.document["status"], "FAILED")
        self.assertEqual(jobs.document["input_file_id"], input_file_id)
        self.assertIn("bad workbook", jobs.document["error_message"])


if __name__ == "__main__":
    unittest.main()
