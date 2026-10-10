"""Startup replay of durable unfinished calculation jobs."""

from .calc_service import CalculationService


def recover_unfinished(database, calculate) -> dict:
    with database.connect() as connection:
        jobs = connection.execute(
            "SELECT copy_id,revision FROM pp_calc_jobs WHERE status='pending' ORDER BY created_at,copy_id,revision"
        ).fetchall()
    service = CalculationService(database)
    completed = failed = 0
    for job in jobs:
        try:
            changes = calculate(job["copy_id"], job["revision"])
            if service.complete(job["copy_id"], job["revision"], changes):
                completed += 1
        except Exception as exc:
            service.failed(job["copy_id"], job["revision"], type(exc).__name__)
            failed += 1
    return {"completed": completed, "failed": failed}
