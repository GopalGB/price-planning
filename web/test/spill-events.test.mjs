import test from "node:test";
import assert from "node:assert/strict";

import { CellStore } from "../src/model.mjs";

const actor = { id: "editor", name: "Synthetic Editor" };
const at = "2026-10-11T00:00:00.000Z";

test("authoritative calc events create absent spill followers with row-shaped metadata", () => {
  const store = new CellStore({ 0: { 1: { v: 1, s: 0, k: "f" } } });

  store.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 7, k: "f" }],
  });

  assert.deepEqual(store.get(0, 1, 1), {
    v: 2,
    s: 7,
    k: "f",
    calcRevision: 1,
  });
});

test("spill expansion and contraction preserve input metadata and null formulas", () => {
  const store = new CellStore({
    0: { 2: { v: 42, s: 8, k: "f" } },
    1: { 1: { v: null, s: 7, k: "i", ver: 3, by: actor, at } },
  });

  store.applyCalc({
    calculated_revision: 1,
    changes: [
      { sheet: 0, r: 1, c: 1, v: 2, s: 7, k: "f" },
      { sheet: 0, r: 0, c: 2, v: null, s: 8, k: "f" },
    ],
  });
  assert.deepEqual(store.get(0, 1, 1), {
    v: 2,
    s: 7,
    k: "f",
    ver: 3,
    by: actor,
    at,
    calcRevision: 1,
  });
  assert.deepEqual(store.get(0, 0, 2), {
    v: null,
    s: 8,
    k: "f",
    calcRevision: 1,
  });

  store.applyCalc({
    calculated_revision: 2,
    changes: [
      { sheet: 0, r: 1, c: 1, v: null, s: 7, k: "i", ver: 3, by: actor, at },
    ],
  });
  assert.deepEqual(store.get(0, 1, 1), {
    v: null,
    s: 7,
    k: "i",
    ver: 3,
    by: actor,
    at,
    calcRevision: 2,
  });
});

test("calc events never overwrite pending or newer input cells", () => {
  const pending = new CellStore({
    1: { 1: { v: null, s: 7, k: "i", ver: 0 } },
  });
  pending.optimistic({
    op_id: "pending-b2",
    cells: [{ sheet: 0, r: 1, c: 1, v: 99 }],
  });
  pending.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 7, k: "f" }],
  });
  assert.deepEqual(pending.get(0, 1, 1), {
    v: 99,
    s: 7,
    k: "i",
    ver: 0,
    pending: "pending-b2",
  });

  const newer = new CellStore({ 1: { 1: { v: null, s: 7, k: "i", ver: 3 } } });
  newer.applyOp({
    op_id: "newer-b2",
    input_revision: 2,
    by: actor,
    at,
    cells: [{ sheet: 0, r: 1, c: 1, v: 99, ver: 4 }],
  });
  newer.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 7, k: "f" }],
  });
  assert.deepEqual(newer.get(0, 1, 1), {
    v: 99,
    s: 7,
    k: "i",
    ver: 4,
    by: actor,
    at,
    lastOpId: "newer-b2",
    pending: undefined,
  });

  const loaded = new CellStore();
  loaded.merge({ 1: { 1: { v: 99, s: 7, k: "i", ver: 4 } } }, 0, 0, 2);
  loaded.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: 1, c: 1, v: 2, s: 7, k: "f" }],
  });
  assert.deepEqual(loaded.get(0, 1, 1), { v: 99, s: 7, k: "i", ver: 4 });
});

test("legacy value-only calc replays retain their prior behavior", () => {
  const store = new CellStore({ 0: { 1: { v: 1, s: 0, k: "f" } } });

  store.applyCalc({
    calculated_revision: 1,
    changes: [
      { sheet: 0, r: 0, c: 1, v: 2 },
      { sheet: 0, r: 1, c: 1, v: 3 },
    ],
  });

  assert.deepEqual(store.get(0, 0, 1), { v: 2, s: 0, k: "f", calcRevision: 1 });
  assert.equal(store.get(0, 1, 1), undefined);
});
