import test from "node:test";
import assert from "node:assert/strict";

import { PlanningGrid } from "../src/grid.mjs";
import { BLOCK_ROWS, CellStore, RowCache } from "../src/model.mjs";

test("released temporary edit rows cannot be repopulated by calc events", () => {
  const store = new CellStore({ 0: { 0: { v: 1, s: 0, k: "i", ver: 0 } } });
  const cache = new RowCache(3);
  cache.put(0, 0, { calculated_revision: 0, rows: {} });
  const cached = { sheet: 0, r: 0, c: 0 };
  const target = { sheet: 0, r: BLOCK_ROWS + 1, c: 0 };
  store.merge({ [target.r]: { 0: { v: 2, s: 0, k: "i", ver: 0 } } }, 0);
  const grid = Object.create(PlanningGrid.prototype);
  Object.assign(grid, { store, cache });
  grid.releaseTemporary([cached, target]);
  assert.equal(store.get(0, target.r, 0), undefined);
  assert.equal(store.get(0, cached.r, cached.c).v, 1);
  store.applyCalc({
    calculated_revision: 1,
    changes: [{ sheet: 0, r: target.r, c: 1, v: 4, s: 0, k: "f" }],
  });
  assert.equal(store.get(0, target.r, 1), undefined);
  assert.equal(store.get(0, 0, 0).v, 1);
  assert.equal(store.blockRevisions.has(`0:${BLOCK_ROWS}`), false);
});
