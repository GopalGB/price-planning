import test from "node:test";
import assert from "node:assert/strict";

import { CellStore } from "../src/model.mjs";

const actor = { id: "editor", name: "Synthetic Editor" };
const at = "2026-10-11T00:00:00.000Z";
const coalesced = {
  type: "calc",
  from_calculated_revision: 1,
  calculated_revision: 3,
  changes: [
    {
      sheet: 0,
      r: 1,
      c: 1,
      v: 2,
      s: 7,
      k: "f",
      input: true,
      input_value: null,
      ver: 2,
      by: actor,
      at,
    },
  ],
};

function spillInputStore() {
  const store = new CellStore();
  store.merge(
    {
      0: { 3: { v: 10, s: 1, k: "i", ver: 0 } },
      1: {
        1: {
          v: 2,
          s: 7,
          k: "f",
          input: true,
          input_value: null,
          ver: 1,
        },
      },
    },
    0,
    1,
    1,
  );
  store.applyOp({
    op_id: "clear-two",
    input_revision: 2,
    cells: [{ sheet: 0, r: 1, c: 1, v: null, ver: 2 }],
  });
  store.applyOp({
    op_id: "edit-d1",
    input_revision: 3,
    cells: [{ sheet: 0, r: 0, c: 3, v: 11, ver: 1 }],
  });
  return store;
}

test("coalesced calc restores a spill input and advances a block at its base", () => {
  const store = spillInputStore();

  store.applyCalc(coalesced);

  assert.deepEqual(store.get(0, 1, 1), {
    v: 2,
    s: 7,
    k: "f",
    input: true,
    input_value: null,
    ver: 2,
    by: actor,
    at,
    lastOpId: "clear-two",
    pending: undefined,
    calcRevision: 3,
  });
  assert.equal(store.blocksCalculated(store.inputRevision), true);
});

test("coalesced intervals advance only blocks that reached the supplied base", () => {
  const store = new CellStore();
  store.merge({ 0: { 0: { v: 1, s: 0, k: "f" } } }, 0, 0);
  store.merge({ 256: { 0: { v: 2, s: 0, k: "f" } } }, 0, 1);
  store.inputRevision = 3;

  store.applyCalc({ ...coalesced, changes: [] });

  assert.equal(store.blockRevisions.get("0:0"), 0);
  assert.equal(store.blockRevisions.get("0:256"), 3);
  assert.equal(store.blocksCalculated(3), false);

  store.applyCalc({ type: "calc", calculated_revision: 1, changes: [] });
  assert.equal(store.blocksCalculated(3), true);

  const refreshed = new CellStore();
  refreshed.merge({ 0: { 0: { v: 1, s: 0, k: "f" } } }, 0, 0);
  refreshed.merge({ 256: { 0: { v: 2, s: 0, k: "f" } } }, 0, 1);
  refreshed.inputRevision = 3;
  refreshed.applyCalc({ ...coalesced, changes: [] });
  assert.equal(refreshed.blocksCalculated(3), false);
  refreshed.registerBlock(0, 0, 3);
  assert.equal(refreshed.blocksCalculated(3), true);

  const legacy = new CellStore();
  legacy.merge({ 0: { 0: { v: 1, s: 0, k: "f" } } }, 0, 1);
  legacy.inputRevision = 3;
  legacy.applyCalc({ type: "calc", calculated_revision: 3, changes: [] });
  assert.equal(legacy.blocksCalculated(3), false);
});
