"""Monotonic persistence for versioned calculation publications."""

import json
from dataclasses import dataclass
from datetime import datetime, timezone

from .values import exactly_equal
from .workbook_service import TemplateCatalog


@dataclass(frozen=True)
class CalculationPublication:
    copy_id: str
    revision: int
    results: list[tuple]
    delta: bool = False
    base_revision: int | None = None
    removals: tuple | list = ()

    @classmethod
    def from_result(cls, copy_id, revision, result):
        return cls(copy_id, revision, result["results"], result.get("delta", False),
                   result.get("base_revision"), result.get("removals", ()))


class PublicationBaseMismatch(Exception):
    """A sparse calculation delta no longer matches the published base."""


@dataclass(frozen=True)
class CalculationDelta:
    copy_id: str
    current: dict
    coordinates: list[tuple]


class CalculationService:
    def __init__(self, database, templates=None):
        self.database = database
        self.templates = (
            templates if templates is None or isinstance(templates, TemplateCatalog)
            else TemplateCatalog(database, templates)
        )

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
            previous_revision = copy["calculated_revision"]
            if publication.delta and publication.base_revision != previous_revision:
                raise PublicationBaseMismatch("calculation publication base changed")
            delta = _published_delta(connection, publication)
            snapshot = self.templates.get(connection, copy["template_id"]) if self.templates else None
            if snapshot:
                delta = _include_touched(connection, delta, (previous_revision, revision))
            _persist_results(connection, publication)
            changes = (_authoritative_changes(connection, snapshot, delta)
                       if snapshot else _legacy_changes(delta))
            payload = {"type": "calc", "from_calculated_revision": previous_revision,
                       "calculated_revision": revision,
                       "changes": changes}
            connection.execute("UPDATE pp_copies SET calculated_revision=?,updated_at=? WHERE id=?",
                               (revision, now, copy_id))
            connection.execute("UPDATE pp_calc_jobs SET status='complete',updated_at=?,last_error=NULL WHERE copy_id=? AND revision=?",
                               (now, copy_id, revision))
            connection.execute("INSERT INTO pp_events(copy_id,event_type,input_revision,payload_json,created_at) VALUES(?,'calc',?,?,?)",
                               (copy_id, revision, _json(payload), now))
            return True

    def published_revision(self, copy_id):
        with self.database.connect() as connection:
            row = connection.execute(
                "SELECT calculated_revision FROM pp_copies WHERE id=?", (copy_id,)
            ).fetchone()
        return row["calculated_revision"] if row is not None else None

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


def _persist_results(connection, publication):
    if not publication.delta:
        _replace_results(connection, publication)
        return
    upserts = {item[:3] for item in publication.results}
    removals = set(publication.removals)
    if upserts & removals:
        raise ValueError("calculation delta cannot upsert and remove the same cell")
    connection.executemany(
        "INSERT INTO pp_results(copy_id,sheet,row,col,value_json,calculated_revision) "
        "VALUES(?,?,?,?,?,?) ON CONFLICT(copy_id,sheet,row,col) DO UPDATE SET "
        "value_json=excluded.value_json,calculated_revision=excluded.calculated_revision",
        [(publication.copy_id, sheet, row, col, _json(value), publication.revision)
         for sheet, row, col, value in publication.results],
    )
    connection.executemany(
        "DELETE FROM pp_results WHERE copy_id=? AND sheet=? AND row=? AND col=?",
        [(publication.copy_id, *coordinate) for coordinate in publication.removals],
    )


def _published_delta(connection, publication):
    current = {item[:3]: item[3] for item in publication.results}
    if publication.delta:
        coordinates = list(current)
        coordinates.extend(coordinate for coordinate in publication.removals
                           if coordinate not in current)
        return CalculationDelta(publication.copy_id, current, coordinates)
    rows = connection.execute(
        "SELECT sheet,row,col,value_json FROM pp_results WHERE copy_id=?", (publication.copy_id,)
    ).fetchall()
    published = {(row["sheet"], row["row"], row["col"]): json.loads(row["value_json"]) for row in rows}
    coordinates = [item[:3] for item in publication.results
                   if item[:3] not in published or not exactly_equal(published[item[:3]], item[3])]
    coordinates.extend(sorted(published.keys() - current.keys()))
    return CalculationDelta(publication.copy_id, current, coordinates)


def _legacy_changes(delta):
    changes = []
    for coordinate in delta.coordinates:
        sheet, row, col = coordinate
        changes.append({"sheet": sheet, "r": row, "c": col,
                        "v": delta.current.get(coordinate)})
    return changes


def _include_touched(connection, delta, revisions):
    previous_revision, revision = revisions
    rows = connection.execute(
        "SELECT DISTINCT h.sheet,h.row,h.col,r.value_json FROM pp_history h JOIN pp_ops o "
        "ON o.copy_id=h.copy_id AND o.op_id=h.op_id AND o.revision=h.revision "
        "LEFT JOIN pp_results r ON r.copy_id=h.copy_id AND r.sheet=h.sheet "
        "AND r.row=h.row AND r.col=h.col "
        "WHERE o.copy_id=? AND o.revision>? AND o.revision<=? "
        "ORDER BY h.sheet,h.row,h.col",
        (delta.copy_id, previous_revision, revision),
    ).fetchall()
    current = dict(delta.current)
    coordinates = list(delta.coordinates)
    known = set(coordinates)
    for row in rows:
        coordinate = row["sheet"], row["row"], row["col"]
        if coordinate not in known and row["value_json"] is not None:
            current[coordinate] = json.loads(row["value_json"])
        if coordinate in current and coordinate not in known:
            coordinates.append(coordinate)
            known.add(coordinate)
    return CalculationDelta(delta.copy_id, current, coordinates)


def _authoritative_changes(connection, snapshot, delta):
    overrides = _override_records(connection, delta.copy_id)
    changes = []
    for coordinate in delta.coordinates:
        sheet, row, col = coordinate
        record = dict(snapshot.cell(sheet, row, col))
        record.pop("formula", None)
        if coordinate in delta.current:
            record.update({"v": delta.current[coordinate], "k": "f"})
            record.update(_spill_input_metadata(snapshot, coordinate, overrides))
        elif record["k"] == "i":
            record.update(overrides.get(coordinate, {"ver": 0}))
        changes.append({"sheet": sheet, "r": row, "c": col, **record})
    return changes


def _spill_input_metadata(snapshot, coordinate, overrides):
    if not snapshot.is_input(*coordinate):
        return {}
    override = overrides.get(coordinate)
    value = override["v"] if override else snapshot.cell(*coordinate)["v"]
    formula = value.get("formula") if isinstance(value, dict) else None
    metadata = {"input": True, "input_value": value if formula is None else formula,
                "ver": override["ver"] if override else 0}
    if formula is not None:
        metadata["input_formula"] = True
    if override:
        metadata.update({key: override[key] for key in ("by", "at")})
    return metadata


def _override_records(connection, copy_id):
    rows = connection.execute(
        "SELECT o.*,u.name FROM pp_overrides o JOIN pp_users u ON u.id=o.updated_by "
        "WHERE o.copy_id=?", (copy_id,),
    ).fetchall()
    return {
        (row["sheet"], row["row"], row["col"]): {
            "v": json.loads(row["value_json"]), "ver": row["version"],
            "by": {"id": row["updated_by"], "name": row["name"]}, "at": row["updated_at"],
        }
        for row in rows
    }


def _now():
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
