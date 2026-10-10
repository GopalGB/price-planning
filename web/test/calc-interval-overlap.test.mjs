import test from "node:test";
import assert from "node:assert/strict";

import { CellStore } from "../src/model.mjs";

test("a complete calc interval advances a block already past its base", () => {
  const store = new CellStore({ 0: { 0: { v: 1, s: 0, k: "f" } } });
  store.registerBlock(0, 0, 2);
  store.applyCalc({
    from_calculated_revision: 1,
    calculated_revision: 3,
    changes: [{ sheet: 0, r: 0, c: 0, v: 3, s: 0, k: "f" }],
  });
  assert.equal(store.get(0, 0, 0).v, 3);
  assert.equal(store.blocksCalculated(3), true);
});
