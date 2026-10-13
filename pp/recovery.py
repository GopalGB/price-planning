"""Startup replay of durable unfinished calculation jobs."""

import inspect

from dataclasses import dataclass

from .calc_service import (
    CalculationPublication, CalculationService, PublicationBaseMismatch,
)


@dataclass(frozen=True)
class RecoveryAttempt:
    copy_id: str
    revision: int
    base_revision: int | None


def recover_unfinished(database, calculate, templates=None) -> dict:
    with database.connect() as connection:
        jobs = connection.execute(
            "SELECT copy_id,revision FROM pp_calc_jobs WHERE status='pending' ORDER BY created_at,copy_id,revision"
        ).fetchall()
    service = CalculationService(database, templates)
    completed = failed = 0
    for job in jobs:
        try:
            attempt = RecoveryAttempt(
                job["copy_id"], job["revision"], service.published_revision(job["copy_id"])
            )
            result = _invoke(calculate, attempt)
            publication = CalculationPublication.from_result(
                job["copy_id"], job["revision"], result
            )
            try:
                published = service.complete(publication)
            except PublicationBaseMismatch:
                result = _invoke(calculate, RecoveryAttempt(job["copy_id"], job["revision"], None))
                publication = CalculationPublication.from_result(
                    job["copy_id"], job["revision"], result
                )
                published = service.complete(publication)
            if published:
                completed += 1
        except Exception as exc:
            service.failed(job["copy_id"], job["revision"], type(exc).__name__)
            failed += 1
    return {"completed": completed, "failed": failed}


def _invoke(calculate, attempt):
    parameters = inspect.signature(calculate).parameters.values()
    positional = [item for item in parameters if item.kind in {
        inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD,
    }]
    variadic = any(item.kind == inspect.Parameter.VAR_POSITIONAL for item in parameters)
    if variadic or len(positional) >= 3:
        return calculate(attempt.copy_id, attempt.revision, attempt.base_revision)
    return calculate(attempt.copy_id, attempt.revision)
