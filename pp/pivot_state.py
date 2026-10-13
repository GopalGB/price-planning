"""Per-copy pivot state: Excel slicer selections and whether the copy's pivots were refreshed (P6.4 decision 2).

Stored as one JSON row per copy so a fork starts from its parent's selections and every calculation
re-applies the same selections after a fresh Excel load.
"""

from __future__ import annotations

import json

EMPTY_STATE = {"slicers": {}, "refreshed": False}


def read_state(connection, copy_id: str) -> dict:
    row = connection.execute("SELECT state_json FROM pp_pivot_state WHERE copy_id=?", (copy_id,)).fetchone()
    if row is None:
        return {"slicers": {}, "refreshed": False}
    state = json.loads(row["state_json"])
    return {"slicers": dict(state.get("slicers", {})), "refreshed": bool(state.get("refreshed", False))}


def write_state(connection, copy_id: str, state: dict, revision: int) -> None:
    value = json.dumps({"slicers": state["slicers"], "refreshed": state["refreshed"]},
                       sort_keys=True, separators=(",", ":"))
    connection.execute(
        "INSERT INTO pp_pivot_state(copy_id,state_json,revision) VALUES(?,?,?) "
        "ON CONFLICT(copy_id) DO UPDATE SET state_json=excluded.state_json, revision=excluded.revision",
        (copy_id, value, revision),
    )


def with_selection(state: dict, cache: str, selected: list[str]) -> dict:
    return {"slicers": {**state["slicers"], cache: list(selected)}, "refreshed": state["refreshed"]}


def with_refreshed(state: dict) -> dict:
    return {"slicers": dict(state["slicers"]), "refreshed": True}
