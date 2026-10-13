"""Background jobs for long workbook transfers, so a browser or proxy timeout never cuts them off.

A job belongs to the session that started it: only that session can read its progress or result.
"""

from __future__ import annotations

from dataclasses import dataclass, field
import logging
import secrets
import threading
import time

from .errors import AppError

LOGGER = logging.getLogger(__name__)
JOB_TTL_SECONDS = 30 * 60
MAX_JOBS = 64


@dataclass
class Job:
    user_id: str
    session_hash: str
    state: str = "running"
    stage: str = "Starting"
    done: int = 0
    total: int = 0
    result: object = None
    error: dict | None = None
    finished_at: float | None = None
    lock: threading.Lock = field(default_factory=threading.Lock)

    def progress(self, stage: str, done: int = 0, total: int = 0) -> None:
        with self.lock:
            self.stage, self.done, self.total = stage, done, total

    def public(self) -> dict:
        with self.lock:
            body = {"state": self.state, "stage": self.stage, "done": self.done, "total": self.total}
            if self.state == "done":
                body["result"] = self.result
            elif self.state == "failed":
                body["error"] = self.error
            return body


class JobStore:
    def __init__(self, ttl: float = JOB_TTL_SECONDS, max_jobs: int = MAX_JOBS):
        self.ttl = ttl
        self.max_jobs = max_jobs
        self._jobs: dict[str, Job] = {}
        self._lock = threading.Lock()

    def start(self, user_id: str, session_hash: str, work, cleanup=None) -> str:
        """Run work(job) on a thread; cleanup() runs after it once the job was accepted."""
        job = Job(user_id, session_hash)
        with self._lock:
            self._purge()
            if len(self._jobs) >= self.max_jobs:
                # Refused before the hand-off, so the caller still owns its resources and releases them
                raise AppError(503, "transfer_busy", "Workbook transfers are busy; try again shortly")
            token = secrets.token_urlsafe(24)
            self._jobs[token] = job
        threading.Thread(target=self._run, args=(job, work, cleanup), name="transfer-job", daemon=True).start()
        return token

    def get(self, token: str, user_id: str, session_hash: str) -> dict:
        with self._lock:
            self._purge()
            job = self._jobs.get(token)
        if job is None or job.user_id != user_id or job.session_hash != session_hash:
            raise AppError(404, "not_found", "This transfer is no longer available")
        return job.public()

    def _run(self, job, work, cleanup):
        try:
            result = work(job)
            state, error = "done", None
        except AppError as exc:
            result, state = None, "failed"
            error = {**exc.extra, "status": exc.status, "code": exc.code, "message": exc.message}
        except Exception:
            LOGGER.exception("Transfer job failed")
            result, state = None, "failed"
            error = {"status": 500, "code": "internal", "message": "The transfer failed on the server"}
        finally:
            if cleanup is not None:
                try:
                    cleanup()
                except Exception:
                    LOGGER.exception("Transfer job cleanup failed")
        with job.lock:
            job.state, job.result, job.error, job.finished_at = state, result, error, time.monotonic()

    def _purge(self):
        now = time.monotonic()
        for token in [token for token, job in self._jobs.items()
                      if job.finished_at is not None and now - job.finished_at > self.ttl]:
            del self._jobs[token]
