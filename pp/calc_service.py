"""Monotonic persistence for complete calculation batches."""

import json
from datetime import datetime, timezone


class CalculationService:
    def __init__(self, database):
        self.database = database

    def complete(self, copy_id: str, revision: int, changes: list[tuple]) -> bool:
        now = _now()
        with self.database.transaction() as connection:
            copy = connection.execute("SELECT * FROM pp_copies WHERE id=?", (copy_id,)).fetchone()
            if copy is None:
                return False
            if revision <= copy["calculated_revision"]:
                connection.execute("UPDATE pp_calc_jobs SET status='superseded',updated_at=? WHERE copy_id=? AND revision=?",
                                   (now, copy_id, revision))
                return False
            if revision > copy["input_revision"]:
                return False
            connection.execute("DELETE FROM pp_results WHERE copy_id=?", (copy_id,))
            for sheet, row, col, value in changes:
                connection.execute(
                    "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) VALUES(?,?,?,?,?,?)",
                    (copy_id, sheet, row, col, _json(value), revision),
                )
            payload = {"type": "calc", "calculated_revision": revision,
                       "changes": [{"sheet": sheet, "r": row, "c": col, "v": value}
                                   for sheet, row, col, value in changes]}
            connection.execute("UPDATE pp_copies SET calculated_revision=?,updated_at=? WHERE id=?",
                               (revision, now, copy_id))
            connection.execute("UPDATE pp_calc_jobs SET status='complete',updated_at=?,last_error=NULL WHERE copy_id=? AND revision=?",
                               (now, copy_id, revision))
            connection.execute("INSERT INTO pp_events(copy_id,event_type,input_revision,payload_json,created_at) VALUES(?,'calc',?,?,?)",
                               (copy_id, revision, _json(payload), now))
            return True

    def failed(self, copy_id: str, revision: int, error: str) -> None:
        with self.database.transaction() as connection:
            connection.execute(
                "UPDATE pp_calc_jobs SET attempts=attempts+1,last_error=?,updated_at=? WHERE copy_id=? AND revision=?",
                (error[:200], _now(), copy_id, revision),
            )


def _json(value):
    return json.dumps(value, separators=(",", ":"), sort_keys=True, allow_nan=False)


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
