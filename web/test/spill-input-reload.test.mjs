import test from "node:test";
import assert from "node:assert/strict";

import { buildOp, CellStore } from "../src/model.mjs";

test("reloaded original input under a spill accepts a normal edit", () => {
  const store = new CellStore({
    1: {
      1: { v: 2, s: 0, k: "f", input: true, input_value: null, ver: 1 },
      2: { v: 3, s: 0, k: "f" },
    },
  });
  let edit;
  assert.doesNotThrow(() => {
    edit = buildOp(store, [{ sheet: 0, r: 1, c: 1, v: 7 }]);
  });
  assert.equal(edit.cells[0].expected_version, 1);
  assert.equal(edit.before["0:1:1"].v, null);
  assert.equal(edit.displayBefore["0:1:1"].v, 2);
  assert.throws(
    () => buildOp(store, [{ sheet: 0, r: 1, c: 2, v: 7 }]),
    /locked or not loaded/,
  );
});
