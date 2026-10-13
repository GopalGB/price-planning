import test from "node:test";
import assert from "node:assert/strict";

import { BLOCK_ROWS, CellStore, RowCache } from "../src/model.mjs";

function calcChange(sheet, r, c = 1) {
  return { sheet, r, c, v: r + 1, s: 7, k: "f" };
}

test("authoritative calc creates spill cells only inside loaded row blocks", () => {
  const store = new CellStore({
    0: {
      0: { v: 2, s: 0, k: "i", ver: 1 },
      1: { v: 1, s: 1, k: "f" },
    },
  });

  store.applyCalc({
    calculated_revision: 1,
    changes: [calcChange(0, 1), calcChange(0, BLOCK_ROWS), calcChange(1, 1)],
  });

  assert.deepEqual(store.get(0, 1, 1), {
    v: 2,
    s: 7,
    k: "f",
    calcRevision: 1,
  });
  assert.equal(store.get(0, BLOCK_ROWS, 1), undefined);
  assert.equal(store.get(1, 1, 1), undefined);
});

test("merging a sparse row marks its whole block as loaded", () => {
  const store = new CellStore();
  store.merge({ [BLOCK_ROWS]: { 0: { v: "marker", s: 0, k: "f" } } }, 0);

  store.applyCalc({
    calculated_revision: 1,
    changes: [calcChange(0, BLOCK_ROWS + 1)],
  });

  assert.deepEqual(store.get(0, BLOCK_ROWS + 1, 1), {
    v: BLOCK_ROWS + 2,
    s: 7,
    k: "f",
    calcRevision: 1,
  });
});

test("row cache eviction prevents calc from repopulating evicted blocks", () => {
  const store = new CellStore();
  const cache = new RowCache(3);
  const starts = [0, BLOCK_ROWS, BLOCK_ROWS * 2, BLOCK_ROWS * 3];

  for (const start of starts) {
    const data = {
      calculated_revision: 0,
      rows: { [start]: { 0: { v: start, s: 0, k: "f" } } },
    };
    cache.put(0, start, data);
    store.registerBlock(0, start, data.calculated_revision);
    for (const evicted of cache.lastEvicted)
      store.evict(evicted.sheet, evicted.start);
    store.merge(data.rows, 0, data.calculated_revision);
  }

  store.applyCalc({
    calculated_revision: 1,
    changes: starts.map((start) => calcChange(0, start + 1)),
  });

  assert.equal(store.get(0, 1, 1), undefined);
  for (const start of starts.slice(1))
    assert.notEqual(store.get(0, start + 1, 1), undefined);
});
