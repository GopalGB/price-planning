"""Monotonic persistence for complete calculation batches."""

import json
from dataclasses import dataclass
from datetime import datetime, timezone

from .values import exactly_equal


@dataclass(frozen=True)
class CalculationPublication:
    copy_id: str
    revision: int
    results: list[tuple]


class CalculationService:
    def __init__(self, database):
        self.database = database

    def complete(self, publication: CalculationPublication) -> bool:
        now = _now()
        copy_id, revision = publication.copy_id, publication.revision
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
            changes = _published_changes(connection, copy_id, publication.results)
            _replace_results(connection, publication)
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


def _replace_results(connection, publication):
    connection.execute("DELETE FROM pp_results WHERE copy_id=?", (publication.copy_id,))
    connection.executemany(
        "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) VALUES(?,?,?,?,?,?)",
        [(publication.copy_id, sheet, row, col, _json(value), publication.revision)
         for sheet, row, col, value in publication.results],
    )


def _published_changes(connection, copy_id, results):
    rows = connection.execute(
        "SELECT sheet,row,col,value_json FROM pp_results WHERE copy_id=?", (copy_id,)
    ).fetchall()
    published = {(row["sheet"], row["row"], row["col"]): json.loads(row["value_json"]) for row in rows}
    current = {item[:3]: item[3] for item in results}
    changes = [item for item in results
               if item[:3] not in published or not exactly_equal(published[item[:3]], item[3])]
    changes.extend((*coordinate, None) for coordinate in sorted(published.keys() - current.keys()))
    return changes


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
