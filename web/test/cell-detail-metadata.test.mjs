import test from "node:test";
import assert from "node:assert/strict";

import { CellStore } from "../src/model.mjs";

test("version-matching detail hydrates a legacy operation missing its timestamp", () => {
  const store = new CellStore({
    1: { 1: { v: 10, s: 0, k: "i", ver: 0 } },
  });
  store.applyOp({
    type: "op",
    op_id: "legacy-op",
    input_revision: 1,
    by: { id: "editor", name: "Editor" },
    cells: [{ sheet: 0, r: 1, c: 1, v: 12, ver: 1 }],
  });
  assert.equal(store.get(0, 1, 1).at, null);

  const applied = store.applyDetail({
    sheet: 0,
    r: 1,
    c: 1,
    v: 12,
    s: 0,
    k: "i",
    ver: 1,
    by: { id: "editor", name: "Editor" },
    at: "2026-10-11T08:00:00.000Z",
  });

  assert.equal(applied, true);
  assert.deepEqual(store.get(0, 1, 1).by, {
    id: "editor",
    name: "Editor",
  });
  assert.equal(store.get(0, 1, 1).at, "2026-10-11T08:00:00.000Z");
});
